using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using Serilog;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Handles migration from legacy flat configuration to hierarchical structure
    /// </summary>
    public class ConfigurationMigrator
    {
        private readonly ILogger _logger;

        public ConfigurationMigrator()
        {
            _logger = Log.ForContext<ConfigurationMigrator>();
        }

        /// <summary>
        /// Migrates from legacy AppConfiguration to hierarchical structure
        /// </summary>
        public HierarchicalAppConfiguration MigrateFromLegacy(AppConfiguration legacyConfig)
        {
            if (legacyConfig == null)
                throw new ArgumentNullException(nameof(legacyConfig));

            _logger.Information("Starting migration from legacy configuration to hierarchical structure");

            var hierarchical = new HierarchicalAppConfiguration();

            try
            {
                // Migrate Application settings
                MigrateApplicationSettings(legacyConfig, hierarchical.Application);

                // Migrate Mode defaults from performance settings
                MigrateModeDefaultSettings(legacyConfig, hierarchical.ModeDefaults);

                // Migrate SaveQuotes mode settings
                MigrateSaveQuotesModeSettings(legacyConfig, hierarchical.Modes);

                // Migrate User preferences
                MigrateUserPreferences(legacyConfig, hierarchical.UserPreferences);

                // Migrate Performance settings
                MigratePerformanceSettings(legacyConfig, hierarchical.Performance);

                // Migrate Display settings
                MigrateDisplaySettings(legacyConfig, hierarchical.Display);

                // Migrate Advanced settings
                MigrateAdvancedSettings(legacyConfig, hierarchical.Advanced);

                // Set metadata
                hierarchical.Metadata.LastModified = DateTime.UtcNow;
                hierarchical.Metadata.MigrationSource = "Legacy JSON Configuration";

                _logger.Information("Configuration migration completed successfully");
                return hierarchical;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to migrate legacy configuration");
                throw;
            }
        }

        private void MigrateApplicationSettings(AppConfiguration legacy, ApplicationSettings application)
        {
            application.Theme = legacy.Theme ?? "Light";
            application.LogLevel = legacy.LogLevel ?? "Information";
            application.DefaultSaveLocation = legacy.DefaultSaveLocation ?? "";
            application.Culture = "en-US"; // Default, can be enhanced later
            application.Language = "English"; // Default, can be enhanced later

            _logger.Debug("Migrated application settings: Theme={Theme}, LogLevel={LogLevel}", 
                application.Theme, application.LogLevel);
        }

        private void MigrateModeDefaultSettings(AppConfiguration legacy, ModeDefaultSettings modeDefaults)
        {
            // Use existing settings as defaults for all modes
            modeDefaults.MaxFileSize = $"{legacy.DocFileSizeLimitMB}MB";
            modeDefaults.ProcessingTimeout = $"{legacy.ConversionTimeoutSeconds}s";
            modeDefaults.MaxConcurrency = legacy.MaxParallelProcessing;
            modeDefaults.EnableValidation = true; // Default
            modeDefaults.EnableProgressReporting = legacy.EnableProgressReporting;

            _logger.Debug("Migrated mode default settings: MaxConcurrency={MaxConcurrency}, ProcessingTimeout={ProcessingTimeout}", 
                modeDefaults.MaxConcurrency, modeDefaults.ProcessingTimeout);
        }

        private void MigrateSaveQuotesModeSettings(AppConfiguration legacy, Dictionary<string, ModeSpecificSettings> modes)
        {
            var saveQuotesSettings = new ModeSpecificSettings
            {
                Enabled = legacy.SaveQuotesMode,
                DisplayName = "Save Quotes",
                Description = "Organize and save quote documents with company names and scope of work",
                Settings = new Dictionary<string, object>
                {
                    ["AutoScanCompanyNames"] = legacy.AutoScanCompanyNames,
                    ["DocFileSizeLimitMB"] = legacy.DocFileSizeLimitMB,
                    ["ScanDocFiles"] = legacy.ScanCompanyNamesForDocFiles,
                    ["ClearScopeAfterProcessing"] = legacy.ClearScopeAfterProcessing,
                    ["ShowRecentScopes"] = legacy.ShowRecentScopes,
                    ["DefaultScope"] = "03-1000" // Default value
                },
                UICustomization = new Dictionary<string, object>
                {
                    ["ShowCompanyDetection"] = true,
                    ["ShowScopeSelector"] = true,
                    ["CompactMode"] = false
                }
            };

            modes["SaveQuotes"] = saveQuotesSettings;

            _logger.Debug("Migrated SaveQuotes mode settings: Enabled={Enabled}, AutoScan={AutoScan}", 
                saveQuotesSettings.Enabled, legacy.AutoScanCompanyNames);
        }

        private void MigrateUserPreferences(AppConfiguration legacy, UserPreferencesSettings userPreferences)
        {
            userPreferences.RecentLocations = new List<string>(legacy.RecentLocations ?? new List<string>());
            userPreferences.MaxRecentLocations = legacy.MaxRecentLocations;

            // Migrate window position
            userPreferences.WindowPosition.Left = legacy.WindowLeft;
            userPreferences.WindowPosition.Top = legacy.WindowTop;
            userPreferences.WindowPosition.Width = legacy.WindowWidth;
            userPreferences.WindowPosition.Height = legacy.WindowHeight;
            userPreferences.WindowPosition.State = legacy.WindowState ?? "Normal";
            userPreferences.WindowPosition.RememberPosition = legacy.RememberWindowPosition;

            // Migrate queue window settings
            userPreferences.QueueWindow.Left = legacy.QueueWindowLeft;
            userPreferences.QueueWindow.Top = legacy.QueueWindowTop;
            userPreferences.QueueWindow.Width = legacy.QueueWindowWidth;
            userPreferences.QueueWindow.Height = legacy.QueueWindowHeight;
            userPreferences.QueueWindow.IsOpen = legacy.QueueWindowIsOpen;
            userPreferences.QueueWindow.RestoreOnStartup = legacy.RestoreQueueWindowOnStartup;

            _logger.Debug("Migrated user preferences: RecentLocations={Count}, WindowSize={Width}x{Height}", 
                userPreferences.RecentLocations.Count, userPreferences.WindowPosition.Width, userPreferences.WindowPosition.Height);
        }

        private void MigratePerformanceSettings(AppConfiguration legacy, PerformanceSettings performance)
        {
            performance.MemoryLimitMB = legacy.MemoryUsageLimitMB;
            performance.EnablePdfCaching = legacy.EnablePdfCaching;
            performance.CacheExpirationMinutes = legacy.PdfCacheExpirationMinutes;
            performance.MaxParallelProcessing = legacy.MaxParallelProcessing;
            performance.ConversionTimeoutSeconds = legacy.ConversionTimeoutSeconds;
            performance.ComTimeoutSeconds = legacy.ComTimeoutSeconds;
            performance.EnableNetworkPathOptimization = legacy.EnableNetworkPathOptimization;

            _logger.Debug("Migrated performance settings: MemoryLimit={MemoryLimit}MB, MaxParallel={MaxParallel}", 
                performance.MemoryLimitMB, performance.MaxParallelProcessing);
        }

        private void MigrateDisplaySettings(AppConfiguration legacy, DisplaySettings display)
        {
            display.OpenFolderAfterProcessing = legacy.OpenFolderAfterProcessing ?? true;
            display.EnableAnimations = legacy.EnableAnimations;
            display.ShowStatusNotifications = legacy.ShowStatusNotifications;
            display.EnableProgressReporting = legacy.EnableProgressReporting;
            display.ShowTooltips = true; // Default value

            _logger.Debug("Migrated display settings: OpenFolder={OpenFolder}, Animations={Animations}", 
                display.OpenFolderAfterProcessing, display.EnableAnimations);
        }

        private void MigrateAdvancedSettings(AppConfiguration legacy, AdvancedSettings advanced)
        {
            advanced.CleanupTempFilesOnExit = legacy.CleanupTempFilesOnExit;
            advanced.EnableDiagnosticMode = legacy.EnableDiagnosticMode;
            advanced.LogFileLocation = legacy.LogFileLocation ?? "";
            advanced.EnableDetailedLogging = false; // Default value
            advanced.DebugMode = false; // Default value

            _logger.Debug("Migrated advanced settings: CleanupTemp={CleanupTemp}, DiagnosticMode={DiagnosticMode}", 
                advanced.CleanupTempFilesOnExit, advanced.EnableDiagnosticMode);
        }

        /// <summary>
        /// Migrates hierarchical configuration back to legacy format for backward compatibility
        /// </summary>
        public AppConfiguration MigrateToLegacy(HierarchicalAppConfiguration hierarchical)
        {
            if (hierarchical == null)
                throw new ArgumentNullException(nameof(hierarchical));

            _logger.Information("Converting hierarchical configuration to legacy format for compatibility");

            var legacy = new AppConfiguration();

            try
            {
                // Application settings
                legacy.Theme = hierarchical.Application.Theme;
                legacy.LogLevel = hierarchical.Application.LogLevel;
                legacy.DefaultSaveLocation = hierarchical.Application.DefaultSaveLocation;

                // Performance settings
                legacy.MaxParallelProcessing = hierarchical.Performance.MaxParallelProcessing;
                legacy.ConversionTimeoutSeconds = hierarchical.Performance.ConversionTimeoutSeconds;
                legacy.MemoryUsageLimitMB = hierarchical.Performance.MemoryLimitMB;
                legacy.EnablePdfCaching = hierarchical.Performance.EnablePdfCaching;
                legacy.PdfCacheExpirationMinutes = hierarchical.Performance.CacheExpirationMinutes;
                legacy.ComTimeoutSeconds = hierarchical.Performance.ComTimeoutSeconds;
                legacy.EnableNetworkPathOptimization = hierarchical.Performance.EnableNetworkPathOptimization;

                // Display settings
                legacy.OpenFolderAfterProcessing = hierarchical.Display.OpenFolderAfterProcessing;
                legacy.EnableAnimations = hierarchical.Display.EnableAnimations;
                legacy.ShowStatusNotifications = hierarchical.Display.ShowStatusNotifications;
                legacy.EnableProgressReporting = hierarchical.Display.EnableProgressReporting;

                // User preferences
                legacy.RecentLocations = new List<string>(hierarchical.UserPreferences.RecentLocations);
                legacy.MaxRecentLocations = hierarchical.UserPreferences.MaxRecentLocations;
                legacy.WindowLeft = hierarchical.UserPreferences.WindowPosition.Left;
                legacy.WindowTop = hierarchical.UserPreferences.WindowPosition.Top;
                legacy.WindowWidth = hierarchical.UserPreferences.WindowPosition.Width;
                legacy.WindowHeight = hierarchical.UserPreferences.WindowPosition.Height;
                legacy.WindowState = hierarchical.UserPreferences.WindowPosition.State;
                legacy.RememberWindowPosition = hierarchical.UserPreferences.WindowPosition.RememberPosition;

                // Queue window
                legacy.QueueWindowLeft = hierarchical.UserPreferences.QueueWindow.Left;
                legacy.QueueWindowTop = hierarchical.UserPreferences.QueueWindow.Top;
                legacy.QueueWindowWidth = hierarchical.UserPreferences.QueueWindow.Width;
                legacy.QueueWindowHeight = hierarchical.UserPreferences.QueueWindow.Height;
                legacy.QueueWindowIsOpen = hierarchical.UserPreferences.QueueWindow.IsOpen;
                legacy.RestoreQueueWindowOnStartup = hierarchical.UserPreferences.QueueWindow.RestoreOnStartup;

                // Advanced settings
                legacy.CleanupTempFilesOnExit = hierarchical.Advanced.CleanupTempFilesOnExit;
                legacy.EnableDiagnosticMode = hierarchical.Advanced.EnableDiagnosticMode;
                legacy.LogFileLocation = hierarchical.Advanced.LogFileLocation;

                // SaveQuotes mode settings
                if (hierarchical.Modes.TryGetValue("SaveQuotes", out var saveQuotesMode))
                {
                    legacy.SaveQuotesMode = saveQuotesMode.Enabled;
                    
                    if (saveQuotesMode.Settings.TryGetValue("AutoScanCompanyNames", out var autoScan))
                        legacy.AutoScanCompanyNames = Convert.ToBoolean(autoScan);
                    
                    if (saveQuotesMode.Settings.TryGetValue("DocFileSizeLimitMB", out var docFileLimit))
                        legacy.DocFileSizeLimitMB = Convert.ToInt32(docFileLimit);
                    
                    if (saveQuotesMode.Settings.TryGetValue("ScanDocFiles", out var scanDocFiles))
                        legacy.ScanCompanyNamesForDocFiles = Convert.ToBoolean(scanDocFiles);
                    
                    if (saveQuotesMode.Settings.TryGetValue("ClearScopeAfterProcessing", out var clearScope))
                        legacy.ClearScopeAfterProcessing = Convert.ToBoolean(clearScope);
                    
                    if (saveQuotesMode.Settings.TryGetValue("ShowRecentScopes", out var showRecent))
                        legacy.ShowRecentScopes = Convert.ToBoolean(showRecent);
                }

                _logger.Information("Successfully converted hierarchical configuration to legacy format");
                return legacy;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to convert hierarchical configuration to legacy format");
                throw;
            }
        }

        /// <summary>
        /// Loads legacy configuration from JSON file
        /// </summary>
        public AppConfiguration? LoadLegacyConfiguration(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            try
            {
                var json = File.ReadAllText(filePath);
                var config = JsonSerializer.Deserialize<AppConfiguration>(json);
                
                _logger.Information("Successfully loaded legacy configuration from {FilePath}", filePath);
                return config;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load legacy configuration from {FilePath}", filePath);
                return null;
            }
        }

        /// <summary>
        /// Creates a backup of the legacy configuration file
        /// </summary>
        public void BackupLegacyConfiguration(string originalPath)
        {
            if (!File.Exists(originalPath))
                return;

            try
            {
                var backupPath = $"{originalPath}.backup.{DateTime.Now:yyyyMMdd_HHmmss}";
                File.Copy(originalPath, backupPath);
                
                _logger.Information("Created backup of legacy configuration at {BackupPath}", backupPath);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to create backup of legacy configuration at {OriginalPath}", originalPath);
            }
        }
    }
} 