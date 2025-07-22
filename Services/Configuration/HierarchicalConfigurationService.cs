using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using Serilog;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Enhanced configuration service with hierarchical YAML support, hot-reload, and migration
    /// </summary>
    public class HierarchicalConfigurationService : IHierarchicalConfigurationService, IDisposable
    {
        private readonly ILogger _logger;
        private readonly string _configDirectory;
        private readonly string _yamlConfigPath;
        private readonly string _legacyJsonPath;
        private readonly ConfigurationMigrator _migrator;
        
        private HierarchicalAppConfiguration _config;
        private FileSystemWatcher? _fileWatcher;
        
        // Serializers
        private readonly ISerializer _yamlSerializer;
        private readonly IDeserializer _yamlDeserializer;
        
        // Debouncing for saves and reload
        private Timer? _saveTimer;
        private Timer? _reloadTimer;
        private readonly SemaphoreSlim _saveSemaphore = new(1, 1);
        private readonly SemaphoreSlim _reloadSemaphore = new(1, 1);
        private bool _saveScheduled = false;
        private bool _reloadScheduled = false;

        // Change notification
        public event EventHandler<ConfigurationChangedEventArgs>? ConfigurationChanged;
        public event EventHandler<ConfigurationErrorEventArgs>? ConfigurationError;

        public HierarchicalAppConfiguration Config => _config;

        public HierarchicalConfigurationService()
        {
            _logger = Log.ForContext<HierarchicalConfigurationService>();
            _migrator = new ConfigurationMigrator();

            // Setup paths
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _configDirectory = Path.Combine(appDataPath, "DocHandler");
            _yamlConfigPath = Path.Combine(_configDirectory, "config.yaml");
            _legacyJsonPath = Path.Combine(_configDirectory, "config.json");

            // Initialize YAML serializers
            _yamlSerializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .WithIndentedSequences()
                .Build();

            _yamlDeserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();

            // Ensure configuration directory exists
            Directory.CreateDirectory(_configDirectory);

            // Load configuration
            _config = LoadConfiguration();

            // Setup file watcher for hot-reload
            SetupFileWatcher();

            _logger.Information("HierarchicalConfigurationService initialized with hot-reload support");
        }

        private HierarchicalAppConfiguration LoadConfiguration()
        {
            try
            {
                // Try to load YAML configuration first
                if (File.Exists(_yamlConfigPath))
                {
                    _logger.Information("Loading YAML configuration from {Path}", _yamlConfigPath);
                    return LoadYamlConfiguration();
                }

                // Fall back to legacy JSON configuration and migrate
                if (File.Exists(_legacyJsonPath))
                {
                    _logger.Information("Legacy JSON configuration found, migrating to YAML");
                    return MigrateLegacyConfiguration();
                }

                // Create default configuration
                _logger.Information("No existing configuration found, creating default hierarchical configuration");
                return CreateDefaultConfiguration();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load configuration, using defaults");
                OnConfigurationError(new ConfigurationErrorEventArgs("Failed to load configuration", ex));
                return CreateDefaultConfiguration();
            }
        }

        private HierarchicalAppConfiguration LoadYamlConfiguration()
        {
            try
            {
                var yaml = File.ReadAllText(_yamlConfigPath);
                var config = _yamlDeserializer.Deserialize<HierarchicalAppConfiguration>(yaml);
                
                if (config != null)
                {
                    _logger.Information("Successfully loaded YAML configuration");
                    return config;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to parse YAML configuration");
                throw;
            }

            throw new InvalidOperationException("Failed to load YAML configuration");
        }

        private HierarchicalAppConfiguration MigrateLegacyConfiguration()
        {
            try
            {
                // Backup legacy configuration
                _migrator.BackupLegacyConfiguration(_legacyJsonPath);

                // Load legacy configuration
                var legacyConfig = _migrator.LoadLegacyConfiguration(_legacyJsonPath);
                if (legacyConfig == null)
                {
                    throw new InvalidOperationException("Failed to load legacy configuration for migration");
                }

                // Migrate to hierarchical structure
                var hierarchicalConfig = _migrator.MigrateFromLegacy(legacyConfig);

                // Save as YAML
                SaveConfigurationImmediate(hierarchicalConfig);

                _logger.Information("Successfully migrated legacy configuration to hierarchical YAML format");
                return hierarchicalConfig;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to migrate legacy configuration");
                throw;
            }
        }

        private HierarchicalAppConfiguration CreateDefaultConfiguration()
        {
            var config = new HierarchicalAppConfiguration();

            // Set default values appropriate for the environment
            config.Application.DefaultSaveLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            // Create default SaveQuotes mode configuration
            config.Modes["SaveQuotes"] = new ModeSpecificSettings
            {
                Enabled = true,
                DisplayName = "Save Quotes",
                Description = "Organize and save quote documents with company names and scope of work",
                Settings = new Dictionary<string, object>
                {
                    ["AutoScanCompanyNames"] = true,
                    ["DocFileSizeLimitMB"] = 10,
                    ["ScanDocFiles"] = false,
                    ["ClearScopeAfterProcessing"] = false,
                    ["ShowRecentScopes"] = false,
                    ["DefaultScope"] = "03-1000"
                },
                UICustomization = new Dictionary<string, object>
                {
                    ["ShowCompanyDetection"] = true,
                    ["ShowScopeSelector"] = true,
                    ["CompactMode"] = false
                }
            };

            config.Metadata.LastModified = DateTime.UtcNow;
            config.Metadata.MigrationSource = "Default Configuration";

            _logger.Information("Created default hierarchical configuration");
            return config;
        }

        private void SetupFileWatcher()
        {
            try
            {
                _fileWatcher = new FileSystemWatcher(_configDirectory)
                {
                    Filter = "config.yaml",
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime,
                    EnableRaisingEvents = true
                };

                _fileWatcher.Changed += OnConfigFileChanged;
                _fileWatcher.Created += OnConfigFileChanged;

                _logger.Debug("File watcher setup for hot-reload: {ConfigPath}", _yamlConfigPath);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to setup file watcher for hot-reload");
            }
        }

        private void OnConfigFileChanged(object sender, FileSystemEventArgs e)
        {
            if (e.FullPath != _yamlConfigPath) return;

            _logger.Debug("Configuration file change detected: {ChangeType}", e.ChangeType);

            // Debounce rapid file changes
            _reloadScheduled = true;
            _reloadTimer?.Dispose();
            
            _reloadTimer = new Timer(async _ =>
            {
                if (_reloadScheduled)
                {
                    _reloadScheduled = false;
                    await ReloadConfigurationAsync();
                }
            }, null, 1000, Timeout.Infinite); // 1 second delay
        }

        private async Task ReloadConfigurationAsync()
        {
            await _reloadSemaphore.WaitAsync();
            try
            {
                _logger.Information("Hot-reloading configuration from file changes");

                var previousConfig = _config;
                var newConfig = LoadYamlConfiguration();

                _config = newConfig;

                // Notify subscribers of configuration change
                OnConfigurationChanged(new ConfigurationChangedEventArgs(previousConfig, newConfig));

                _logger.Information("Configuration hot-reload completed successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to hot-reload configuration");
                OnConfigurationError(new ConfigurationErrorEventArgs("Hot-reload failed", ex));
            }
            finally
            {
                _reloadSemaphore.Release();
            }
        }

        public async Task SaveConfigurationAsync()
        {
            await SaveConfigurationDebounced();
        }

        private async Task SaveConfigurationDebounced()
        {
            _saveScheduled = true;

            // Cancel any existing timer
            _saveTimer?.Dispose();

            // Schedule save after 500ms of inactivity
            _saveTimer = new Timer(async _ =>
            {
                if (_saveScheduled)
                {
                    _saveScheduled = false;
                    await SaveConfigurationImmediate(_config);
                }
            }, null, 500, Timeout.Infinite);
        }

        private async Task SaveConfigurationImmediate(HierarchicalAppConfiguration config)
        {
            await _saveSemaphore.WaitAsync();
            try
            {
                // Temporarily disable file watcher to prevent reload during save
                if (_fileWatcher != null)
                    _fileWatcher.EnableRaisingEvents = false;

                // Update metadata
                config.Metadata.LastModified = DateTime.UtcNow;

                // Serialize to YAML
                var yaml = _yamlSerializer.Serialize(config);
                await File.WriteAllTextAsync(_yamlConfigPath, yaml);

                _logger.Debug("Configuration saved to YAML: {Path}", _yamlConfigPath);

                // Also save legacy JSON for backward compatibility
                await SaveLegacyCompatibilityFile(config);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save configuration");
                OnConfigurationError(new ConfigurationErrorEventArgs("Save failed", ex));
                throw;
            }
            finally
            {
                // Re-enable file watcher
                if (_fileWatcher != null)
                {
                    await Task.Delay(100); // Brief delay to avoid immediate reload
                    _fileWatcher.EnableRaisingEvents = true;
                }
                
                _saveSemaphore.Release();
            }
        }

        private async Task SaveLegacyCompatibilityFile(HierarchicalAppConfiguration hierarchical)
        {
            try
            {
                var legacyConfig = _migrator.MigrateToLegacy(hierarchical);
                var options = new JsonSerializerOptions { WriteIndented = true };
                var json = JsonSerializer.Serialize(legacyConfig, options);
                await File.WriteAllTextAsync(_legacyJsonPath, json);

                _logger.Debug("Legacy compatibility file updated: {Path}", _legacyJsonPath);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to update legacy compatibility file");
            }
        }

        public void UpdateConfiguration(Action<HierarchicalAppConfiguration> updateAction)
        {
            if (updateAction == null)
                throw new ArgumentNullException(nameof(updateAction));

            updateAction(_config);
            _ = SaveConfigurationAsync();
        }

        public T GetModeConfiguration<T>(string modeName) where T : class, new()
        {
            if (string.IsNullOrWhiteSpace(modeName))
                throw new ArgumentException("Mode name cannot be null or empty", nameof(modeName));

            if (!_config.Modes.TryGetValue(modeName, out var modeSettings))
            {
                _logger.Warning("Mode configuration not found: {ModeName}", modeName);
                return new T();
            }

            try
            {
                // Convert settings dictionary to typed object
                var json = JsonSerializer.Serialize(modeSettings.Settings);
                var typedConfig = JsonSerializer.Deserialize<T>(json);
                return typedConfig ?? new T();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to deserialize mode configuration for {ModeName}", modeName);
                return new T();
            }
        }

        public void UpdateModeConfiguration<T>(string modeName, T configuration) where T : class
        {
            if (string.IsNullOrWhiteSpace(modeName))
                throw new ArgumentException("Mode name cannot be null or empty", nameof(modeName));

            if (configuration == null)
                throw new ArgumentNullException(nameof(configuration));

            if (!_config.Modes.ContainsKey(modeName))
            {
                _config.Modes[modeName] = new ModeSpecificSettings
                {
                    Enabled = true,
                    DisplayName = modeName,
                    Description = $"Configuration for {modeName} mode"
                };
            }

            try
            {
                // Convert typed object to settings dictionary
                var json = JsonSerializer.Serialize(configuration);
                var settings = JsonSerializer.Deserialize<Dictionary<string, object>>(json);
                
                if (settings != null)
                {
                    _config.Modes[modeName].Settings = settings;
                    _ = SaveConfigurationAsync();
                    
                    _logger.Debug("Updated mode configuration for {ModeName}", modeName);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to update mode configuration for {ModeName}", modeName);
                throw;
            }
        }

        public async Task<string> ExportConfigurationAsync(string? filePath = null)
        {
            try
            {
                filePath ??= Path.Combine(_configDirectory, $"config_export_{DateTime.Now:yyyyMMdd_HHmmss}.yaml");
                
                var yaml = _yamlSerializer.Serialize(_config);
                await File.WriteAllTextAsync(filePath, yaml);
                
                _logger.Information("Configuration exported to {FilePath}", filePath);
                return filePath;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to export configuration to {FilePath}", filePath);
                throw;
            }
        }

        public async Task ImportConfigurationAsync(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Configuration file not found: {filePath}");

            try
            {
                var yaml = await File.ReadAllTextAsync(filePath);
                var importedConfig = _yamlDeserializer.Deserialize<HierarchicalAppConfiguration>(yaml);
                
                if (importedConfig != null)
                {
                    var previousConfig = _config;
                    _config = importedConfig;
                    
                    await SaveConfigurationImmediate(_config);
                    
                    OnConfigurationChanged(new ConfigurationChangedEventArgs(previousConfig, _config));
                    
                    _logger.Information("Configuration imported from {FilePath}", filePath);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to import configuration from {FilePath}", filePath);
                throw;
            }
        }

        public void ResetToDefaults()
        {
            var previousConfig = _config;
            _config = CreateDefaultConfiguration();
            
            _ = SaveConfigurationAsync();
            
            OnConfigurationChanged(new ConfigurationChangedEventArgs(previousConfig, _config));
            
            _logger.Information("Configuration reset to default values");
        }

        private void OnConfigurationChanged(ConfigurationChangedEventArgs args)
        {
            try
            {
                ConfigurationChanged?.Invoke(this, args);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error occurred in configuration change handler");
            }
        }

        private void OnConfigurationError(ConfigurationErrorEventArgs args)
        {
            try
            {
                ConfigurationError?.Invoke(this, args);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error occurred in configuration error handler");
            }
        }

        public void Dispose()
        {
            _fileWatcher?.Dispose();
            _saveTimer?.Dispose();
            _reloadTimer?.Dispose();
            _saveSemaphore.Dispose();
            _reloadSemaphore.Dispose();
        }
    }

    /// <summary>
    /// Event arguments for configuration changes
    /// </summary>
    public class ConfigurationChangedEventArgs : EventArgs
    {
        public HierarchicalAppConfiguration PreviousConfiguration { get; }
        public HierarchicalAppConfiguration NewConfiguration { get; }
        public DateTime ChangeTime { get; }

        public ConfigurationChangedEventArgs(HierarchicalAppConfiguration previousConfig, HierarchicalAppConfiguration newConfig)
        {
            PreviousConfiguration = previousConfig;
            NewConfiguration = newConfig;
            ChangeTime = DateTime.UtcNow;
        }
    }

    /// <summary>
    /// Event arguments for configuration errors
    /// </summary>
    public class ConfigurationErrorEventArgs : EventArgs
    {
        public string Message { get; }
        public Exception? Exception { get; }
        public DateTime ErrorTime { get; }

        public ConfigurationErrorEventArgs(string message, Exception? exception = null)
        {
            Message = message;
            Exception = exception;
            ErrorTime = DateTime.UtcNow;
        }
    }
} 