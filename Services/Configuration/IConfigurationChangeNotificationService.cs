using System;
using System.Threading.Tasks;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Interface for configuration change notification service
    /// </summary>
    public interface IConfigurationChangeNotificationService
    {
        /// <summary>
        /// Subscribe to changes in a specific configuration section (async)
        /// </summary>
        void SubscribeToSection<T>(string sectionName, Func<T, Task> handler) where T : class;

        /// <summary>
        /// Subscribe to changes in a specific configuration section (sync)
        /// </summary>
        void SubscribeToSection<T>(string sectionName, Action<T> handler) where T : class;

        /// <summary>
        /// Subscribe to global configuration changes (async)
        /// </summary>
        void SubscribeToGlobalChanges(Func<HierarchicalAppConfiguration, HierarchicalAppConfiguration, Task> handler);

        /// <summary>
        /// Subscribe to global configuration changes (sync)
        /// </summary>
        void SubscribeToGlobalChanges(Action<HierarchicalAppConfiguration, HierarchicalAppConfiguration> handler);

        /// <summary>
        /// Subscribe to performance settings changes
        /// </summary>
        void SubscribeToPerformanceChanges(Action<PerformanceSettings> handler);

        /// <summary>
        /// Subscribe to application settings changes
        /// </summary>
        void SubscribeToApplicationChanges(Action<ApplicationSettings> handler);

        /// <summary>
        /// Subscribe to mode-specific settings changes
        /// </summary>
        void SubscribeToModeChanges(string modeName, Action<ModeSpecificSettings> handler);

        /// <summary>
        /// Subscribe to display settings changes
        /// </summary>
        void SubscribeToDisplayChanges(Action<DisplaySettings> handler);

        /// <summary>
        /// Subscribe to user preferences changes
        /// </summary>
        void SubscribeToUserPreferencesChanges(Action<UserPreferencesSettings> handler);
    }
} 