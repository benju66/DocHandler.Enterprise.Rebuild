using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Serilog;
using System.Linq; // Added for .Take() and .Distinct()
using DocHandler.Services.Configuration; // Added for SaveQuotesConfiguration

namespace DocHandler.Services
{
    public class ConfigurationService : IConfigurationService, IDisposable
    {
        private readonly ILogger _logger;
        private readonly string _configPath;
        private AppConfiguration _config;
        
        // Thread-safe configuration access
        private readonly ReaderWriterLockSlim _configLock = new(LockRecursionPolicy.SupportsRecursion);
        private readonly SemaphoreSlim _saveSemaphore = new(1, 1);
        private readonly SemaphoreSlim _loadSemaphore = new(1, 1);
        
        // Debouncing for configuration saves
        private Timer? _saveTimer;
        private volatile bool _saveScheduled = false;
        private volatile bool _disposed = false;
        private readonly object _timerLock = new object();
        
        public AppConfiguration Config 
        { 
            get
            {
                _configLock.EnterReadLock();
                try
                {
                    return _config;
                }
                finally
                {
                    _configLock.ExitReadLock();
                }
            }
        }
        
        public ConfigurationService()
        {
            _logger = Log.ForContext<ConfigurationService>();
            
            // Store config in AppData
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "DocHandler");
            
            try
            {
                Directory.CreateDirectory(appFolder);
                _configPath = Path.Combine(appFolder, "config.json");
                _config = LoadConfigurationInternal();
                
                _logger.Information("Configuration service initialized with path: {Path}", _configPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to initialize configuration service");
                _config = CreateDefaultConfiguration();
            }
        }
        
        private AppConfiguration LoadConfigurationInternal()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var json = File.ReadAllText(_configPath);
                    var config = JsonSerializer.Deserialize<AppConfiguration>(json, GetJsonOptions());
                    
                    if (config != null)
                    {
                        // Validate and sanitize loaded configuration
                        ValidateAndSanitizeConfiguration(config);
                        
                        _logger.Information("Configuration loaded from {Path}", _configPath);
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load configuration from {Path}", _configPath);
                
                // Try to backup corrupted config
                TryBackupCorruptedConfig();
            }
            
            // Return default configuration
            _logger.Information("Using default configuration");
            return CreateDefaultConfiguration();
        }
        
        private void TryBackupCorruptedConfig()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var backupPath = $"{_configPath}.backup.{DateTime.UtcNow:yyyyMMdd_HHmmss}";
                    File.Copy(_configPath, backupPath);
                    _logger.Information("Backed up corrupted config to {BackupPath}", backupPath);
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to backup corrupted configuration");
            }
        }
        
        /// <summary>
        /// Loads configuration asynchronously with proper error handling
        /// </summary>
        public async Task<AppConfiguration> LoadConfigurationAsync()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ConfigurationService));
            
            await _loadSemaphore.WaitAsync();
            try
            {
                return await Task.Run(() =>
                {
                    if (File.Exists(_configPath))
                    {
                        var json = File.ReadAllText(_configPath);
                        var config = JsonSerializer.Deserialize<AppConfiguration>(json, GetJsonOptions());
                        
                        if (config != null)
                        {
                            ValidateAndSanitizeConfiguration(config);
                            _logger.Information("Configuration loaded asynchronously from {Path}", _configPath);
                            
                            // Update current configuration thread-safely
                            _configLock.EnterWriteLock();
                            try
                            {
                                _config = config;
                            }
                            finally
                            {
                                _configLock.ExitWriteLock();
                            }
                            
                            return config;
                        }
                    }
                    
                    _logger.Information("No valid configuration found, using current or default");
                    return Config; // Return current config if load fails
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load configuration asynchronously");
                return Config; // Return current config on error
            }
            finally
            {
                _loadSemaphore.Release();
            }
        }
        
        /// <summary>
        /// Updates configuration with thread safety
        /// </summary>
        public void UpdateConfiguration(Action<AppConfiguration> updateAction)
        {
            if (_disposed) return;
            
            _configLock.EnterWriteLock();
            try
            {
                updateAction(_config);
                ValidateAndSanitizeConfiguration(_config);
                
                // Schedule save after update
                ScheduleSaveConfiguration();
            }
            finally
            {
                _configLock.ExitWriteLock();
            }
        }
        
        /// <summary>
        /// Updates configuration asynchronously with thread safety
        /// </summary>
        public async Task UpdateConfigurationAsync(Func<AppConfiguration, Task> updateAction)
        {
            if (_disposed) return;
            
            // Create a copy for async operations to avoid holding locks too long
            AppConfiguration configCopy;
            _configLock.EnterReadLock();
            try
            {
                configCopy = JsonSerializer.Deserialize<AppConfiguration>(
                    JsonSerializer.Serialize(_config, GetJsonOptions()), 
                    GetJsonOptions()) ?? CreateDefaultConfiguration();
            }
            finally
            {
                _configLock.ExitReadLock();
            }
            
            // Perform async update on copy
            await updateAction(configCopy);
            ValidateAndSanitizeConfiguration(configCopy);
            
            // Apply updated copy back to main config
            _configLock.EnterWriteLock();
            try
            {
                _config = configCopy;
                ScheduleSaveConfiguration();
            }
            finally
            {
                _configLock.ExitWriteLock();
            }
        }
        
        private void ValidateAndSanitizeConfiguration(AppConfiguration config)
        {
            // Ensure required properties have valid values
            config.RecentLocations ??= new List<string>();
            config.DefaultSaveLocation ??= Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            
            // Limit recent locations to prevent unbounded growth
            if (config.RecentLocations.Count > config.MaxRecentLocations)
            {
                config.RecentLocations = config.RecentLocations
                    .Take(config.MaxRecentLocations)
                    .ToList();
            }
            
            // Validate paths exist or are accessible
            if (!Directory.Exists(config.DefaultSaveLocation))
            {
                _logger.Warning("Default save location does not exist, using Desktop: {Path}", config.DefaultSaveLocation);
                config.DefaultSaveLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }
            
            // Remove non-existent recent locations
            config.RecentLocations = config.RecentLocations
                .Where(Directory.Exists)
                .Distinct()
                .ToList();
            
            // Validate numeric limits
            if (config.MaxRecentLocations <= 0)
                config.MaxRecentLocations = 10;
            
            if (config.MaxParallelProcessing <= 0)
                config.MaxParallelProcessing = Math.Min(Environment.ProcessorCount, 4);
        }
        
        private JsonSerializerOptions GetJsonOptions()
        {
            return new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNameCaseInsensitive = true,
                AllowTrailingCommas = true,
                ReadCommentHandling = JsonCommentHandling.Skip
            };
        }
        
        private AppConfiguration CreateDefaultConfiguration()
        {
            return new AppConfiguration
            {
                DefaultSaveLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                RecentLocations = new List<string>(),
                MaxRecentLocations = 10,
                SaveQuotesMode = false,
                OpenFolderAfterProcessing = true,
                MaxParallelProcessing = Math.Min(Environment.ProcessorCount, 4),
                EnablePdfCaching = true,
                PdfCacheExpirationMinutes = 30,
                
                // Security settings
                MaxFileSizeMB = 50,
                EnableSecurityScanning = true,
                
                // Performance settings
                MemoryUsageLimitMB = 500,
                ConversionTimeoutSeconds = 30,
                
                // UI settings
                Theme = "Light",
                ShowAdvancedOptions = false,
                
                // Telemetry settings
                EnableTelemetry = true,
                TelemetryLevel = "Normal",
                
                // SaveQuotes Mode Configuration
                SaveQuotes = new SaveQuotesConfiguration()
            };
        }

        // Implement interface method
        public async Task SaveConfigurationAsync()
        {
            await SaveConfiguration();
        }

        public async Task<bool> SaveConfiguration()
        {
            if (_disposed) return false;
            
            await _saveSemaphore.WaitAsync();
            try
            {
                AppConfiguration configToSave;
                _configLock.EnterReadLock();
                try
                {
                    configToSave = _config;
                }
                finally
                {
                    _configLock.ExitReadLock();
                }
                
                var json = JsonSerializer.Serialize(configToSave, GetJsonOptions());
                
                // Write to temporary file first, then rename (atomic operation)
                var tempPath = _configPath + ".tmp";
                await File.WriteAllTextAsync(tempPath, json);
                
                // Atomic replace
                if (File.Exists(_configPath))
                    File.Delete(_configPath);
                File.Move(tempPath, _configPath);
                
                _logger.Debug("Configuration saved to {Path}", _configPath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save configuration to {Path}", _configPath);
                return false;
            }
            finally
            {
                _saveSemaphore.Release();
            }
        }

        // Missing methods needed by MainViewModel
        public void UpdateDefaultSaveLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location) || _disposed) return;
            
            _configLock.EnterWriteLock();
            try
            {
                _config.DefaultSaveLocation = location;
                
                // Add to recent locations if it's a directory
                if (Directory.Exists(location) && !_config.RecentLocations.Contains(location))
                {
                    _config.RecentLocations.Insert(0, location);
                    if (_config.RecentLocations.Count > _config.MaxRecentLocations)
                    {
                        _config.RecentLocations.RemoveAt(_config.RecentLocations.Count - 1);
                    }
                }
                
                _logger.Information("Default save location updated to: {Location}", location);
            }
            finally
            {
                _configLock.ExitWriteLock();
            }
        }

        public void UpdateWindowPosition(double left, double top, double width, double height, string state = "Normal")
        {
            if (_disposed) return;
            
            _configLock.EnterWriteLock();
            try
            {
                _config.WindowLeft = left;
                _config.WindowTop = top;
                _config.WindowWidth = width;
                _config.WindowHeight = height;
                _config.WindowState = state ?? "Normal";
                
                _logger.Debug("Window position updated: {Left},{Top} {Width}x{Height} State:{State}", 
                    left, top, width, height, state);
            }
            finally
            {
                _configLock.ExitWriteLock();
            }
        }

        public void UpdateTheme(string theme)
        {
            if (string.IsNullOrWhiteSpace(theme) || _disposed) return;
            
            _configLock.EnterWriteLock();
            try
            {
                _config.Theme = theme;
                _logger.Information("Theme updated to: {Theme}", theme);
            }
            finally
            {
                _configLock.ExitWriteLock();
            }
        }

        public AppConfiguration GetDefaultConfiguration()
        {
            return CreateDefaultConfiguration();
        }
        
        private void ScheduleSaveConfiguration()
        {
            if (_disposed) return;
            
            lock (_timerLock)
            {
                if (_saveScheduled) return;
                
                _saveScheduled = true;
                
                // Dispose existing timer if it exists
                _saveTimer?.Dispose();
                
                // Create new timer for debounced save (save after 2 seconds of inactivity)
                _saveTimer = new Timer(async _ =>
                {
                    lock (_timerLock)
                    {
                        _saveScheduled = false;
                    }
                    
                    try
                    {
                        await SaveConfiguration();
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error during scheduled configuration save");
                    }
                }, null, TimeSpan.FromSeconds(2), Timeout.InfiniteTimeSpan);
            }
        }
        
        public void AddRecentLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location) || !Directory.Exists(location))
                return;
            
            UpdateConfiguration(config =>
            {
                // Remove if already exists
                config.RecentLocations.Remove(location);
                
                // Add to beginning
                config.RecentLocations.Insert(0, location);
                
                // Trim to max size
                if (config.RecentLocations.Count > config.MaxRecentLocations)
                {
                    config.RecentLocations = config.RecentLocations
                        .Take(config.MaxRecentLocations)
                        .ToList();
                }
            });
        }
        
        public void RemoveRecentLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location))
                return;
            
            UpdateConfiguration(config =>
            {
                config.RecentLocations.Remove(location);
            });
        }
        
        public void ClearRecentLocations()
        {
            UpdateConfiguration(config =>
            {
                config.RecentLocations.Clear();
            });
        }
        
        /// <summary>
        /// Forces immediate save of configuration
        /// </summary>
        public async Task<bool> SaveConfigurationImmediately()
        {
            if (_disposed) return false;
            
            // Cancel any pending scheduled save
            lock (_timerLock)
            {
                _saveTimer?.Dispose();
                _saveTimer = null;
                _saveScheduled = false;
            }
            
            return await SaveConfiguration();
        }
        
        public void Dispose()
        {
            if (_disposed) return;
            
            _logger.Information("Disposing configuration service");
            _disposed = true;
            
            // Save final configuration synchronously
            try
            {
                SaveConfiguration().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save configuration during disposal");
            }
            
            // Dispose resources
            lock (_timerLock)
            {
                _saveTimer?.Dispose();
                _saveTimer = null;
            }
            
            _configLock?.Dispose();
            _saveSemaphore?.Dispose();
            _loadSemaphore?.Dispose();
            
            _logger.Information("Configuration service disposed");
        }
    }
    
    public class AppConfiguration
    {
        public string DefaultSaveLocation { get; set; } = "";
        public List<string> RecentLocations { get; set; } = new();
        public int MaxRecentLocations { get; set; } = 10;
        public string Theme { get; set; } = "Light";
        public bool RememberWindowPosition { get; set; } = true;
        public double WindowLeft { get; set; }
        public double WindowTop { get; set; }
        public double WindowWidth { get; set; }
        public double WindowHeight { get; set; }
        public string WindowState { get; set; } = "Normal";
        public bool? OpenFolderAfterProcessing { get; set; } = true;
        public bool SaveQuotesMode { get; set; } = true;
        public bool ShowRecentScopes { get; set; } = false;
        public bool AutoScanCompanyNames { get; set; } = true;
        public bool ScanCompanyNamesForDocFiles { get; set; } = false;
        public int DocFileSizeLimitMB { get; set; } = 10;
        public bool ClearScopeAfterProcessing { get; set; } = false;
        
        // Queue Window State
        public double? QueueWindowLeft { get; set; }
        public double? QueueWindowTop { get; set; }
        public double? QueueWindowWidth { get; set; } = 600;
        public double? QueueWindowHeight { get; set; } = 400;
        public bool QueueWindowIsOpen { get; set; } = false;
        public bool RestoreQueueWindowOnStartup { get; set; } = true;
        
        // Performance Settings
        public int MaxParallelProcessing { get; set; } = 3;
        public int ConversionTimeoutSeconds { get; set; } = 30;
        public bool EnablePdfCaching { get; set; } = true;
        public int PdfCacheExpirationMinutes { get; set; } = 30;
        public bool EnableProgressReporting { get; set; } = true;
        public int MemoryUsageLimitMB { get; set; } = 500;
        
        // Additional Display Settings
        public bool EnableAnimations { get; set; } = true;
        public bool ShowStatusNotifications { get; set; } = true;
        
        // Additional Advanced Settings
        public bool CleanupTempFilesOnExit { get; set; } = true;
        public bool EnableDiagnosticMode { get; set; } = false;
        public int ComTimeoutSeconds { get; set; } = 30;
        public bool EnableNetworkPathOptimization { get; set; } = true;
        public string LogLevel { get; set; } = "Information";
        public string LogFileLocation { get; set; } = "";
        
        // New properties for enhanced configuration
        public bool ConvertOfficeToPdf { get; set; } = true;
        public bool EnablePdfCache { get; set; } = true;
        public int CacheExpirationMinutes { get; set; } = 30;
        public bool EnableFileValidation { get; set; } = true;
        public bool EnableCompanyDetection { get; set; } = true;
        public bool EnableScopeDetection { get; set; } = true;
        public bool AutoScanDocuments { get; set; } = false;
        public List<string> AllowedFileExtensions { get; set; } = new() { ".docx", ".doc", ".xlsx", ".xls", ".pdf", ".txt", ".rtf" };
        public int MaxFileSizeMB { get; set; } = 50;
        public bool EnableSecurityScanning { get; set; } = true;
        public int ProcessingTimeoutMinutes { get; set; } = 30;
        public int MaxConcurrentConversions { get; set; } = Environment.ProcessorCount;
        public string Language { get; set; } = "en-US";
        public bool ShowAdvancedOptions { get; set; } = false;
        public bool EnableTelemetry { get; set; } = true;
        public string TelemetryLevel { get; set; } = "Normal";
        
        // SaveQuotes Mode Configuration
        public SaveQuotesConfiguration SaveQuotes { get; set; } = new SaveQuotesConfiguration();
    }
}