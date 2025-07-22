using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Text.Json;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using Serilog;
using System.Collections.Generic; // Added missing import for List

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Service for exporting and importing configuration with validation and backup functionality
    /// </summary>
    public class ConfigurationExportImportService : IConfigurationExportImportService
    {
        private readonly ILogger _logger;
        private readonly IHierarchicalConfigurationService _configService;
        private readonly ISerializer _yamlSerializer;
        private readonly IDeserializer _yamlDeserializer;

        public ConfigurationExportImportService(IHierarchicalConfigurationService configService)
        {
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _logger = Log.ForContext<ConfigurationExportImportService>();

            _yamlSerializer = new SerializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .WithIndentedSequences()
                .Build();

            _yamlDeserializer = new DeserializerBuilder()
                .WithNamingConvention(CamelCaseNamingConvention.Instance)
                .IgnoreUnmatchedProperties()
                .Build();

            _logger.Information("Configuration export/import service initialized");
        }

        /// <summary>
        /// Export current configuration to YAML file
        /// </summary>
        public async Task<string> ExportConfigurationAsync(string? filePath = null, ExportOptions? options = null)
        {
            try
            {
                options ??= new ExportOptions();
                
                // Generate filename if not provided
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var fileName = $"DocHandler_Config_Export_{timestamp}.yaml";
                    filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName);
                }

                var config = _configService.Config;
                
                // Apply export options
                var exportConfig = options.IncludeSensitiveData ? config : SanitizeConfiguration(config);

                // Add export metadata
                exportConfig.Metadata.LastModified = DateTime.UtcNow;
                exportConfig.Metadata.CreatedBy = "DocHandler Export Service";
                exportConfig.Metadata.MigrationSource = "Configuration Export";

                // Serialize to YAML
                var yaml = _yamlSerializer.Serialize(exportConfig);
                
                // Add export header with instructions
                var exportContent = BuildExportHeader(options) + yaml;
                
                await File.WriteAllTextAsync(filePath, exportContent);

                _logger.Information("Configuration exported successfully to {FilePath}", filePath);
                return filePath;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to export configuration to {FilePath}", filePath);
                throw;
            }
        }

        /// <summary>
        /// Import configuration from YAML file with validation
        /// </summary>
        public async Task<ImportResult> ImportConfigurationAsync(string filePath, ImportOptions? options = null)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Configuration file not found: {filePath}");

            options ??= new ImportOptions();
            
            try
            {
                _logger.Information("Starting configuration import from {FilePath}", filePath);

                // Read and parse the file
                var yamlContent = await File.ReadAllTextAsync(filePath);
                
                // Remove export header if present
                yamlContent = RemoveExportHeader(yamlContent);
                
                var importedConfig = _yamlDeserializer.Deserialize<HierarchicalAppConfiguration>(yamlContent);
                
                if (importedConfig == null)
                    throw new InvalidOperationException("Failed to deserialize configuration from file");

                // Validate the imported configuration
                var validationResult = ValidateConfiguration(importedConfig);
                if (!validationResult.IsValid)
                {
                    return new ImportResult
                    {
                        Success = false,
                        ErrorMessage = $"Configuration validation failed: {validationResult.ErrorMessage}",
                        ValidationErrors = validationResult.ValidationErrors
                    };
                }

                // Create backup if requested
                string? backupPath = null;
                if (options.CreateBackup)
                {
                    backupPath = await CreateConfigurationBackup();
                }

                // Apply the configuration
                if (options.MergeWithExisting)
                {
                    await MergeConfiguration(importedConfig);
                }
                else
                {
                    await _configService.ImportConfigurationAsync(filePath);
                }

                _logger.Information("Configuration imported successfully from {FilePath}", filePath);

                return new ImportResult
                {
                    Success = true,
                    BackupPath = backupPath,
                    ImportedSettingsCount = CountConfigurationSettings(importedConfig),
                    Message = "Configuration imported successfully"
                };
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to import configuration from {FilePath}", filePath);
                return new ImportResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        /// <summary>
        /// Create a backup of the current configuration
        /// </summary>
        public async Task<string> CreateConfigurationBackup()
        {
            try
            {
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var backupFileName = $"DocHandler_Config_Backup_{timestamp}.yaml";
                var backupPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                    "DocHandler", 
                    "Backups", 
                    backupFileName);

                // Ensure backup directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(backupPath)!);

                var backupFilePath = await _configService.ExportConfigurationAsync(backupPath);
                
                _logger.Information("Configuration backup created at {BackupPath}", backupFilePath);
                return backupFilePath;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to create configuration backup");
                throw;
            }
        }

        /// <summary>
        /// Validate configuration structure and values
        /// </summary>
        public ConfigurationValidationResult ValidateConfiguration(HierarchicalAppConfiguration config)
        {
            var result = new ConfigurationValidationResult { IsValid = true };

            try
            {
                // Validate application settings
                if (string.IsNullOrWhiteSpace(config.Application.Theme))
                {
                    result.AddError("Application.Theme cannot be empty");
                }

                if (string.IsNullOrWhiteSpace(config.Application.LogLevel))
                {
                    result.AddError("Application.LogLevel cannot be empty");
                }

                // Validate performance settings
                if (config.Performance.MemoryLimitMB < 100 || config.Performance.MemoryLimitMB > 8192)
                {
                    result.AddError("Performance.MemoryLimitMB must be between 100 and 8192 MB");
                }

                if (config.Performance.MaxParallelProcessing < 1 || config.Performance.MaxParallelProcessing > 20)
                {
                    result.AddError("Performance.MaxParallelProcessing must be between 1 and 20");
                }

                // Validate mode settings
                foreach (var mode in config.Modes)
                {
                    if (string.IsNullOrWhiteSpace(mode.Value.DisplayName))
                    {
                        result.AddError($"Mode '{mode.Key}' must have a DisplayName");
                    }
                }

                // Validate user preferences
                if (config.UserPreferences.MaxRecentLocations < 5 || config.UserPreferences.MaxRecentLocations > 50)
                {
                    result.AddWarning("UserPreferences.MaxRecentLocations should be between 5 and 50");
                }

                result.IsValid = result.ValidationErrors.Count == 0;
                
                _logger.Debug("Configuration validation completed: {IsValid}, Errors: {ErrorCount}, Warnings: {WarningCount}", 
                    result.IsValid, result.ValidationErrors.Count, result.ValidationWarnings.Count);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during configuration validation");
                result.IsValid = false;
                result.ErrorMessage = $"Validation error: {ex.Message}";
                return result;
            }
        }

        private HierarchicalAppConfiguration SanitizeConfiguration(HierarchicalAppConfiguration config)
        {
            // Create a copy and remove sensitive data
            var sanitized = JsonSerializer.Deserialize<HierarchicalAppConfiguration>(
                JsonSerializer.Serialize(config));

            // Remove or mask sensitive information
            sanitized!.UserPreferences.RecentLocations.Clear();
            
            // Clear any API keys or sensitive settings in mode configurations
            foreach (var mode in sanitized.Modes)
            {
                if (mode.Value.Settings.ContainsKey("ApiKey"))
                    mode.Value.Settings["ApiKey"] = "***REDACTED***";
                
                if (mode.Value.Settings.ContainsKey("Password"))
                    mode.Value.Settings["Password"] = "***REDACTED***";
            }

            return sanitized;
        }

        private string BuildExportHeader(ExportOptions options)
        {
            var header = $@"# DocHandler Enterprise Configuration Export
# Generated: {DateTime.UtcNow:yyyy-MM-dd HH:mm:ss} UTC
# Export Options: 
#   - Include Sensitive Data: {options.IncludeSensitiveData}
#   - Export Format: YAML
# 
# INSTRUCTIONS:
# 1. Review the configuration before importing
# 2. Backup your current configuration before importing
# 3. Test in a non-production environment first
# 4. Validate all paths and settings for your environment
#
# NOTE: This configuration was exported for DocHandler Enterprise
# Make sure the target system has compatible versions and features
#
---
";
            return header;
        }

        private string RemoveExportHeader(string yamlContent)
        {
            // Remove everything before the first "---" line
            var yamlStart = yamlContent.IndexOf("---");
            if (yamlStart >= 0)
            {
                return yamlContent.Substring(yamlStart + 3).Trim();
            }
            return yamlContent;
        }

        private async Task MergeConfiguration(HierarchicalAppConfiguration importedConfig)
        {
            // For now, implement a simple merge strategy
            // In production, you might want more sophisticated merging logic
            
            _configService.UpdateConfiguration(currentConfig =>
            {
                // Merge application settings
                if (!string.IsNullOrWhiteSpace(importedConfig.Application.Theme))
                    currentConfig.Application.Theme = importedConfig.Application.Theme;
                
                if (!string.IsNullOrWhiteSpace(importedConfig.Application.LogLevel))
                    currentConfig.Application.LogLevel = importedConfig.Application.LogLevel;

                // Merge performance settings (with validation)
                if (importedConfig.Performance.MemoryLimitMB > 0)
                    currentConfig.Performance.MemoryLimitMB = importedConfig.Performance.MemoryLimitMB;

                // Merge mode settings
                foreach (var importedMode in importedConfig.Modes)
                {
                    currentConfig.Modes[importedMode.Key] = importedMode.Value;
                }
            });

            await _configService.SaveConfigurationAsync();
        }

        private int CountConfigurationSettings(HierarchicalAppConfiguration config)
        {
            int count = 0;
            
            // Count main sections
            count += 7; // Application, Performance, Display, UserPreferences, ModeDefaults, Advanced, Metadata
            
            // Count mode settings
            foreach (var mode in config.Modes)
            {
                count += mode.Value.Settings.Count;
            }

            return count;
        }
    }

    /// <summary>
    /// Options for configuration export
    /// </summary>
    public class ExportOptions
    {
        public bool IncludeSensitiveData { get; set; } = false;
        public bool PrettyFormat { get; set; } = true;
        public bool IncludeComments { get; set; } = true;
    }

    /// <summary>
    /// Options for configuration import
    /// </summary>
    public class ImportOptions
    {
        public bool CreateBackup { get; set; } = true;
        public bool MergeWithExisting { get; set; } = false;
        public bool ValidateBeforeImport { get; set; } = true;
    }

    /// <summary>
    /// Result of configuration import operation
    /// </summary>
    public class ImportResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public string? BackupPath { get; set; }
        public int ImportedSettingsCount { get; set; }
        public string? Message { get; set; }
        public List<string> ValidationErrors { get; set; } = new List<string>();
    }

    /// <summary>
    /// Result of configuration validation
    /// </summary>
    public class ConfigurationValidationResult
    {
        public bool IsValid { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> ValidationErrors { get; set; } = new List<string>();
        public List<string> ValidationWarnings { get; set; } = new List<string>();

        public void AddError(string error)
        {
            ValidationErrors.Add(error);
            IsValid = false;
        }

        public void AddWarning(string warning)
        {
            ValidationWarnings.Add(warning);
        }
    }
} 