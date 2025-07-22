using System;
using System.Threading.Tasks;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Interface for hierarchical configuration service with hot-reload support
    /// </summary>
    public interface IHierarchicalConfigurationService
    {
        /// <summary>
        /// Current hierarchical configuration
        /// </summary>
        HierarchicalAppConfiguration Config { get; }

        /// <summary>
        /// Event fired when configuration changes (including hot-reload)
        /// </summary>
        event EventHandler<ConfigurationChangedEventArgs>? ConfigurationChanged;

        /// <summary>
        /// Event fired when configuration errors occur
        /// </summary>
        event EventHandler<ConfigurationErrorEventArgs>? ConfigurationError;

        /// <summary>
        /// Save configuration asynchronously
        /// </summary>
        Task SaveConfigurationAsync();

        /// <summary>
        /// Update configuration with action
        /// </summary>
        void UpdateConfiguration(Action<HierarchicalAppConfiguration> updateAction);

        /// <summary>
        /// Get typed configuration for a specific mode
        /// </summary>
        T GetModeConfiguration<T>(string modeName) where T : class, new();

        /// <summary>
        /// Update typed configuration for a specific mode
        /// </summary>
        void UpdateModeConfiguration<T>(string modeName, T configuration) where T : class;

        /// <summary>
        /// Export configuration to file
        /// </summary>
        Task<string> ExportConfigurationAsync(string? filePath = null);

        /// <summary>
        /// Import configuration from file
        /// </summary>
        Task ImportConfigurationAsync(string filePath);

        /// <summary>
        /// Reset configuration to default values
        /// </summary>
        void ResetToDefaults();
    }
} 