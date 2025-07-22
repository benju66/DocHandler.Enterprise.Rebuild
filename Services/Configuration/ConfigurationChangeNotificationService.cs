using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Service that manages configuration change notifications and service subscriptions
    /// </summary>
    public class ConfigurationChangeNotificationService : IConfigurationChangeNotificationService, IDisposable
    {
        private readonly ILogger _logger;
        private readonly IHierarchicalConfigurationService _configService;
        
        // Subscribers organized by configuration section
        private readonly Dictionary<string, List<Func<object, Task>>> _sectionSubscribers = new();
        private readonly Dictionary<string, List<Action<object>>> _syncSectionSubscribers = new();
        
        // Service-wide subscribers
        private readonly List<Func<HierarchicalAppConfiguration, HierarchicalAppConfiguration, Task>> _globalSubscribers = new();
        private readonly List<Action<HierarchicalAppConfiguration, HierarchicalAppConfiguration>> _syncGlobalSubscribers = new();
        
        private bool _disposed = false;

        public ConfigurationChangeNotificationService(IHierarchicalConfigurationService configService)
        {
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _logger = Log.ForContext<ConfigurationChangeNotificationService>();
            
            // Subscribe to configuration changes
            _configService.ConfigurationChanged += OnConfigurationChanged;
            
            _logger.Information("Configuration change notification service initialized");
        }

        /// <summary>
        /// Subscribe to changes in a specific configuration section
        /// </summary>
        public void SubscribeToSection<T>(string sectionName, Func<T, Task> handler) where T : class
        {
            if (string.IsNullOrWhiteSpace(sectionName))
                throw new ArgumentException("Section name cannot be null or empty", nameof(sectionName));

            if (handler == null)
                throw new ArgumentNullException(nameof(handler));

            if (!_sectionSubscribers.ContainsKey(sectionName))
            {
                _sectionSubscribers[sectionName] = new List<Func<object, Task>>();
            }

            _sectionSubscribers[sectionName].Add(async obj =>
            {
                if (obj is T typedObj)
                {
                    await handler(typedObj);
                }
            });

            _logger.Debug("Added async subscription for section: {SectionName}", sectionName);
        }

        /// <summary>
        /// Subscribe to changes in a specific configuration section (synchronous)
        /// </summary>
        public void SubscribeToSection<T>(string sectionName, Action<T> handler) where T : class
        {
            if (string.IsNullOrWhiteSpace(sectionName))
                throw new ArgumentException("Section name cannot be null or empty", nameof(sectionName));

            if (handler == null)
                throw new ArgumentNullException(nameof(handler));

            if (!_syncSectionSubscribers.ContainsKey(sectionName))
            {
                _syncSectionSubscribers[sectionName] = new List<Action<object>>();
            }

            _syncSectionSubscribers[sectionName].Add(obj =>
            {
                if (obj is T typedObj)
                {
                    handler(typedObj);
                }
            });

            _logger.Debug("Added sync subscription for section: {SectionName}", sectionName);
        }

        /// <summary>
        /// Subscribe to global configuration changes
        /// </summary>
        public void SubscribeToGlobalChanges(Func<HierarchicalAppConfiguration, HierarchicalAppConfiguration, Task> handler)
        {
            if (handler == null)
                throw new ArgumentNullException(nameof(handler));

            _globalSubscribers.Add(handler);
            _logger.Debug("Added global async configuration change subscriber");
        }

        /// <summary>
        /// Subscribe to global configuration changes (synchronous)
        /// </summary>
        public void SubscribeToGlobalChanges(Action<HierarchicalAppConfiguration, HierarchicalAppConfiguration> handler)
        {
            if (handler == null)
                throw new ArgumentNullException(nameof(handler));

            _syncGlobalSubscribers.Add(handler);
            _logger.Debug("Added global sync configuration change subscriber");
        }

        /// <summary>
        /// Subscribe to performance settings changes
        /// </summary>
        public void SubscribeToPerformanceChanges(Action<PerformanceSettings> handler)
        {
            SubscribeToSection("Performance", handler);
        }

        /// <summary>
        /// Subscribe to application settings changes
        /// </summary>
        public void SubscribeToApplicationChanges(Action<ApplicationSettings> handler)
        {
            SubscribeToSection("Application", handler);
        }

        /// <summary>
        /// Subscribe to mode-specific settings changes
        /// </summary>
        public void SubscribeToModeChanges(string modeName, Action<ModeSpecificSettings> handler)
        {
            if (string.IsNullOrWhiteSpace(modeName))
                throw new ArgumentException("Mode name cannot be null or empty", nameof(modeName));

            SubscribeToSection($"Mode_{modeName}", handler);
        }

        /// <summary>
        /// Subscribe to display settings changes
        /// </summary>
        public void SubscribeToDisplayChanges(Action<DisplaySettings> handler)
        {
            SubscribeToSection("Display", handler);
        }

        /// <summary>
        /// Subscribe to user preferences changes
        /// </summary>
        public void SubscribeToUserPreferencesChanges(Action<UserPreferencesSettings> handler)
        {
            SubscribeToSection("UserPreferences", handler);
        }

        private async void OnConfigurationChanged(object? sender, ConfigurationChangedEventArgs e)
        {
            try
            {
                _logger.Information("Processing configuration change notifications");

                var previousConfig = e.PreviousConfiguration;
                var newConfig = e.NewConfiguration;

                // Notify global subscribers
                await NotifyGlobalSubscribers(previousConfig, newConfig);

                // Notify section-specific subscribers
                await NotifySectionSubscribers(previousConfig, newConfig);

                _logger.Information("Configuration change notifications completed successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error occurred while processing configuration change notifications");
            }
        }

        private async Task NotifyGlobalSubscribers(HierarchicalAppConfiguration previous, HierarchicalAppConfiguration current)
        {
            // Notify async global subscribers
            foreach (var subscriber in _globalSubscribers)
            {
                try
                {
                    await subscriber(previous, current);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error in global async configuration change subscriber");
                }
            }

            // Notify sync global subscribers
            foreach (var subscriber in _syncGlobalSubscribers)
            {
                try
                {
                    subscriber(previous, current);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error in global sync configuration change subscriber");
                }
            }
        }

        private async Task NotifySectionSubscribers(HierarchicalAppConfiguration previous, HierarchicalAppConfiguration current)
        {
            // Check each section for changes and notify relevant subscribers
            await NotifySectionIfChanged("Application", previous.Application, current.Application);
            await NotifySectionIfChanged("Performance", previous.Performance, current.Performance);
            await NotifySectionIfChanged("Display", previous.Display, current.Display);
            await NotifySectionIfChanged("UserPreferences", previous.UserPreferences, current.UserPreferences);
            await NotifySectionIfChanged("ModeDefaults", previous.ModeDefaults, current.ModeDefaults);

            // Check mode-specific changes
            await NotifyModeChanges(previous.Modes, current.Modes);
        }

        private async Task NotifySectionIfChanged<T>(string sectionName, T previousValue, T currentValue) where T : class
        {
            // Simple equality check - in production, you might want more sophisticated change detection
            if (!ReferenceEquals(previousValue, currentValue))
            {
                await NotifySection(sectionName, currentValue);
            }
        }

        private async Task NotifyModeChanges(Dictionary<string, ModeSpecificSettings> previousModes, Dictionary<string, ModeSpecificSettings> currentModes)
        {
            // Check each mode for changes
            foreach (var kvp in currentModes)
            {
                var modeName = kvp.Key;
                var currentModeSettings = kvp.Value;
                
                if (!previousModes.TryGetValue(modeName, out var previousModeSettings) || 
                    !ReferenceEquals(previousModeSettings, currentModeSettings))
                {
                    await NotifySection($"Mode_{modeName}", currentModeSettings);
                }
            }
        }

        private async Task NotifySection(string sectionName, object sectionValue)
        {
            // Notify async subscribers
            if (_sectionSubscribers.TryGetValue(sectionName, out var asyncSubscribers))
            {
                foreach (var subscriber in asyncSubscribers)
                {
                    try
                    {
                        await subscriber(sectionValue);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error in async section subscriber for {SectionName}", sectionName);
                    }
                }
            }

            // Notify sync subscribers
            if (_syncSectionSubscribers.TryGetValue(sectionName, out var syncSubscribers))
            {
                foreach (var subscriber in syncSubscribers)
                {
                    try
                    {
                        subscriber(sectionValue);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error in sync section subscriber for {SectionName}", sectionName);
                    }
                }
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _configService.ConfigurationChanged -= OnConfigurationChanged;
                
                _sectionSubscribers.Clear();
                _syncSectionSubscribers.Clear();
                _globalSubscribers.Clear();
                _syncGlobalSubscribers.Clear();
                
                _disposed = true;
                _logger.Information("Configuration change notification service disposed");
            }
        }
    }
} 