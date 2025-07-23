using System;
using System.Collections.Generic;
using YamlDotNet.Serialization;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Hierarchical application configuration structure for DocHandler Enterprise
    /// </summary>
    public class HierarchicalAppConfiguration
    {
        /// <summary>
        /// Global application settings
        /// </summary>
        [YamlMember(Alias = "Application")]
        public ApplicationSettings Application { get; set; } = new ApplicationSettings();

        /// <summary>
        /// Default settings for all modes
        /// </summary>
        [YamlMember(Alias = "ModeDefaults")]
        public ModeDefaultSettings ModeDefaults { get; set; } = new ModeDefaultSettings();

        /// <summary>
        /// Mode-specific configuration settings
        /// </summary>
        [YamlMember(Alias = "Modes")]
        public Dictionary<string, ModeSpecificSettings> Modes { get; set; } = new Dictionary<string, ModeSpecificSettings>();

        /// <summary>
        /// User preferences and personalization
        /// </summary>
        [YamlMember(Alias = "UserPreferences")]
        public UserPreferencesSettings UserPreferences { get; set; } = new UserPreferencesSettings();

        /// <summary>
        /// Performance and optimization settings
        /// </summary>
        [YamlMember(Alias = "Performance")]
        public PerformanceSettings Performance { get; set; } = new PerformanceSettings();

        /// <summary>
        /// Display and UI settings
        /// </summary>
        [YamlMember(Alias = "Display")]
        public DisplaySettings Display { get; set; } = new DisplaySettings();

        /// <summary>
        /// Advanced configuration settings
        /// </summary>
        [YamlMember(Alias = "Advanced")]
        public AdvancedSettings Advanced { get; set; } = new AdvancedSettings();

        /// <summary>
        /// Telemetry and Application Insights settings
        /// </summary>
        [YamlMember(Alias = "Telemetry")]
        public TelemetrySettings Telemetry { get; set; } = new TelemetrySettings();

        /// <summary>
        /// Configuration metadata
        /// </summary>
        [YamlMember(Alias = "Metadata")]
        public ConfigurationMetadata Metadata { get; set; } = new ConfigurationMetadata();
    }

    /// <summary>
    /// Global application settings
    /// </summary>
    public class ApplicationSettings
    {
        [YamlMember(Alias = "Theme")]
        public string Theme { get; set; } = "Light";

        [YamlMember(Alias = "LogLevel")]
        public string LogLevel { get; set; } = "Information";

        [YamlMember(Alias = "DefaultSaveLocation")]
        public string DefaultSaveLocation { get; set; } = "";

        [YamlMember(Alias = "Culture")]
        public string Culture { get; set; } = "en-US";

        [YamlMember(Alias = "Language")]
        public string Language { get; set; } = "English";
    }

    /// <summary>
    /// Default settings applied to all modes unless overridden
    /// </summary>
    public class ModeDefaultSettings
    {
        [YamlMember(Alias = "MaxFileSize")]
        public string MaxFileSize { get; set; } = "50MB";

        [YamlMember(Alias = "ProcessingTimeout")]
        public string ProcessingTimeout { get; set; } = "300s";

        [YamlMember(Alias = "MaxConcurrency")]
        public int MaxConcurrency { get; set; } = 5;

        [YamlMember(Alias = "EnableValidation")]
        public bool EnableValidation { get; set; } = true;

        [YamlMember(Alias = "EnableProgressReporting")]
        public bool EnableProgressReporting { get; set; } = true;
    }

    /// <summary>
    /// Mode-specific configuration settings
    /// </summary>
    public class ModeSpecificSettings
    {
        [YamlMember(Alias = "Enabled")]
        public bool Enabled { get; set; } = true;

        [YamlMember(Alias = "DisplayName")]
        public string DisplayName { get; set; } = "";

        [YamlMember(Alias = "Description")]
        public string Description { get; set; } = "";

        [YamlMember(Alias = "Settings")]
        public Dictionary<string, object> Settings { get; set; } = new Dictionary<string, object>();

        [YamlMember(Alias = "UICustomization")]
        public Dictionary<string, object> UICustomization { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// User preferences and personalization settings
    /// </summary>
    public class UserPreferencesSettings
    {
        [YamlMember(Alias = "RecentLocations")]
        public List<string> RecentLocations { get; set; } = new List<string>();

        [YamlMember(Alias = "MaxRecentLocations")]
        public int MaxRecentLocations { get; set; } = 10;

        [YamlMember(Alias = "WindowPosition")]
        public WindowPositionSettings WindowPosition { get; set; } = new WindowPositionSettings();

        [YamlMember(Alias = "QueueWindow")]
        public QueueWindowSettings QueueWindow { get; set; } = new QueueWindowSettings();

        [YamlMember(Alias = "LastUsedSettings")]
        public Dictionary<string, object> LastUsedSettings { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// Main window position settings
    /// </summary>
    public class WindowPositionSettings
    {
        [YamlMember(Alias = "Left")]
        public double Left { get; set; } = 100;

        [YamlMember(Alias = "Top")]
        public double Top { get; set; } = 100;

        [YamlMember(Alias = "Width")]
        public double Width { get; set; } = 800;

        [YamlMember(Alias = "Height")]
        public double Height { get; set; } = 600;

        [YamlMember(Alias = "State")]
        public string State { get; set; } = "Normal";

        [YamlMember(Alias = "RememberPosition")]
        public bool RememberPosition { get; set; } = true;
    }

    /// <summary>
    /// Queue window settings
    /// </summary>
    public class QueueWindowSettings
    {
        [YamlMember(Alias = "Left")]
        public double? Left { get; set; }

        [YamlMember(Alias = "Top")]
        public double? Top { get; set; }

        [YamlMember(Alias = "Width")]
        public double? Width { get; set; } = 600;

        [YamlMember(Alias = "Height")]
        public double? Height { get; set; } = 400;

        [YamlMember(Alias = "IsOpen")]
        public bool IsOpen { get; set; } = false;

        [YamlMember(Alias = "RestoreOnStartup")]
        public bool RestoreOnStartup { get; set; } = true;
    }

    /// <summary>
    /// Performance and optimization settings
    /// </summary>
    public class PerformanceSettings
    {
        [YamlMember(Alias = "MemoryLimitMB")]
        public int MemoryLimitMB { get; set; } = 500;

        [YamlMember(Alias = "EnablePdfCaching")]
        public bool EnablePdfCaching { get; set; } = true;

        [YamlMember(Alias = "CacheExpirationMinutes")]
        public int CacheExpirationMinutes { get; set; } = 30;

        [YamlMember(Alias = "MaxParallelProcessing")]
        public int MaxParallelProcessing { get; set; } = 3;

        [YamlMember(Alias = "ConversionTimeoutSeconds")]
        public int ConversionTimeoutSeconds { get; set; } = 30;

        [YamlMember(Alias = "ComTimeoutSeconds")]
        public int ComTimeoutSeconds { get; set; } = 30;

        [YamlMember(Alias = "EnableNetworkPathOptimization")]
        public bool EnableNetworkPathOptimization { get; set; } = true;
    }

    /// <summary>
    /// Display and UI settings
    /// </summary>
    public class DisplaySettings
    {
        [YamlMember(Alias = "OpenFolderAfterProcessing")]
        public bool OpenFolderAfterProcessing { get; set; } = true;

        [YamlMember(Alias = "EnableAnimations")]
        public bool EnableAnimations { get; set; } = true;

        [YamlMember(Alias = "ShowStatusNotifications")]
        public bool ShowStatusNotifications { get; set; } = true;

        [YamlMember(Alias = "EnableProgressReporting")]
        public bool EnableProgressReporting { get; set; } = true;

        [YamlMember(Alias = "ShowTooltips")]
        public bool ShowTooltips { get; set; } = true;
    }

    /// <summary>
    /// Advanced configuration settings
    /// </summary>
    public class AdvancedSettings
    {
        [YamlMember(Alias = "CleanupTempFilesOnExit")]
        public bool CleanupTempFilesOnExit { get; set; } = true;

        [YamlMember(Alias = "EnableDiagnosticMode")]
        public bool EnableDiagnosticMode { get; set; } = false;

        [YamlMember(Alias = "LogFileLocation")]
        public string LogFileLocation { get; set; } = "";

        [YamlMember(Alias = "EnableDetailedLogging")]
        public bool EnableDetailedLogging { get; set; } = false;

        [YamlMember(Alias = "DebugMode")]
        public bool DebugMode { get; set; } = false;
    }

    /// <summary>
    /// Configuration file metadata
    /// </summary>
    public class ConfigurationMetadata
    {
        [YamlMember(Alias = "Version")]
        public string Version { get; set; } = "2.0.0";

        [YamlMember(Alias = "LastModified")]
        public DateTime LastModified { get; set; } = DateTime.UtcNow;

        [YamlMember(Alias = "CreatedBy")]
        public string CreatedBy { get; set; } = "DocHandler Enterprise";

        [YamlMember(Alias = "SchemaVersion")]
        public string SchemaVersion { get; set; } = "1.0";

        [YamlMember(Alias = "MigrationSource")]
        public string? MigrationSource { get; set; }
    }

    /// <summary>
    /// Telemetry and Application Insights configuration settings
    /// </summary>
    public class TelemetrySettings
    {
        /// <summary>
        /// Enable Application Insights telemetry
        /// </summary>
        [YamlMember(Alias = "EnableApplicationInsights")]
        public bool EnableApplicationInsights { get; set; } = false;

        /// <summary>
        /// Application Insights connection string
        /// </summary>
        [YamlMember(Alias = "ApplicationInsightsConnectionString")]
        public string ApplicationInsightsConnectionString { get; set; } = "";

        /// <summary>
        /// Application Insights instrumentation key (legacy)
        /// </summary>
        [YamlMember(Alias = "ApplicationInsightsInstrumentationKey")]
        public string ApplicationInsightsInstrumentationKey { get; set; } = "";

        /// <summary>
        /// Enable custom business metrics
        /// </summary>
        [YamlMember(Alias = "EnableBusinessMetrics")]
        public bool EnableBusinessMetrics { get; set; } = true;

        /// <summary>
        /// Enable performance counters
        /// </summary>
        [YamlMember(Alias = "EnablePerformanceCounters")]
        public bool EnablePerformanceCounters { get; set; } = true;

        /// <summary>
        /// Enable user interaction tracking
        /// </summary>
        [YamlMember(Alias = "EnableUserInteractionTracking")]
        public bool EnableUserInteractionTracking { get; set; } = true;

        /// <summary>
        /// Enable dependency tracking
        /// </summary>
        [YamlMember(Alias = "EnableDependencyTracking")]
        public bool EnableDependencyTracking { get; set; } = true;

        /// <summary>
        /// Telemetry sampling percentage (0-100)
        /// </summary>
        [YamlMember(Alias = "SamplingPercentage")]
        public double SamplingPercentage { get; set; } = 100.0;

        /// <summary>
        /// Enable local debugging mode (logs to console)
        /// </summary>
        [YamlMember(Alias = "EnableLocalDebugging")]
        public bool EnableLocalDebugging { get; set; } = true;

        /// <summary>
        /// Flush telemetry interval in seconds
        /// </summary>
        [YamlMember(Alias = "FlushIntervalSeconds")]
        public int FlushIntervalSeconds { get; set; } = 30;
    }
} 