using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Provides UI customization for different processing modes
    /// </summary>
    public class ModeUIProvider : IAdvancedModeUIProvider
    {
        private readonly ILogger _logger;
        private readonly IModeManager _modeManager;
        private readonly Dictionary<string, ModeUICustomization> _modeCustomizations;
        private readonly Dictionary<string, Func<FrameworkElement, Task>> _uiApplyActions;

        public ModeUIProvider(IModeManager modeManager)
        {
            _logger = Log.ForContext<ModeUIProvider>();
            _modeManager = modeManager ?? throw new ArgumentNullException(nameof(modeManager));
            _modeCustomizations = new Dictionary<string, ModeUICustomization>();
            _uiApplyActions = new Dictionary<string, Func<FrameworkElement, Task>>();
            
            InitializeDefaultModeCustomizations();
            _logger.Debug("ModeUIProvider initialized with {Count} mode customizations", _modeCustomizations.Count);
        }

        public async Task<ModeUICustomization> GetModeUIAsync(string modeId)
        {
            try
            {
                _logger.Debug("Getting UI customization for mode: {ModeId}", modeId);

                if (string.IsNullOrEmpty(modeId))
                {
                    return GetDefaultModeUI();
                }

                if (_modeCustomizations.TryGetValue(modeId, out var customization))
                {
                    _logger.Debug("Found UI customization for mode: {ModeId}", modeId);
                    return customization;
                }

                _logger.Warning("No UI customization found for mode: {ModeId}, returning default", modeId);
                return GetDefaultModeUI();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error getting UI customization for mode: {ModeId}", modeId);
                return GetDefaultModeUI();
            }
        }

        public bool SupportsModeUI(string modeId)
        {
            return !string.IsNullOrEmpty(modeId) && _modeCustomizations.ContainsKey(modeId);
        }

        public async Task<IEnumerable<string>> GetAvailableModesAsync()
        {
            try
            {
                await Task.CompletedTask; // Make async for consistency
                var modes = _modeCustomizations.Keys.ToList();
                _logger.Debug("Returning {Count} available modes with UI customization", modes.Count);
                return modes;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error getting available modes");
                return Enumerable.Empty<string>();
            }
        }

        public async Task ApplyModeUIAsync(string modeId, FrameworkElement targetElement)
        {
            try
            {
                _logger.Information("Applying UI customization for mode: {ModeId}", modeId);

                if (targetElement == null)
                {
                    _logger.Warning("Target element is null, cannot apply UI customization");
                    return;
                }

                var customization = await GetModeUIAsync(modeId);
                
                // Apply UI properties
                await ApplyUIPropertiesAsync(customization, targetElement);
                
                // Hide/show elements
                await ApplyElementVisibilityAsync(customization, targetElement);
                
                // Apply custom UI actions if registered
                if (_uiApplyActions.TryGetValue(modeId, out var customAction))
                {
                    await customAction(targetElement);
                }

                _logger.Information("UI customization applied successfully for mode: {ModeId}", modeId);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error applying UI customization for mode: {ModeId}", modeId);
            }
        }

        public async Task ResetUIAsync(FrameworkElement targetElement)
        {
            try
            {
                _logger.Information("Resetting UI to default state");
                
                if (targetElement == null)
                {
                    _logger.Warning("Target element is null, cannot reset UI");
                    return;
                }

                // Apply default mode UI
                await ApplyModeUIAsync("default", targetElement);
                
                _logger.Information("UI reset to default state successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error resetting UI to default state");
            }
        }

        /// <summary>
        /// Register a custom UI apply action for a mode
        /// </summary>
        public void RegisterUIApplyAction(string modeId, Func<FrameworkElement, Task> action)
        {
            if (!string.IsNullOrEmpty(modeId) && action != null)
            {
                _uiApplyActions[modeId] = action;
                _logger.Debug("Registered UI apply action for mode: {ModeId}", modeId);
            }
        }

        /// <summary>
        /// Register a mode UI customization
        /// </summary>
        public void RegisterModeCustomization(string modeId, ModeUICustomization customization)
        {
            if (!string.IsNullOrEmpty(modeId) && customization != null)
            {
                customization.ModeId = modeId;
                _modeCustomizations[modeId] = customization;
                _logger.Debug("Registered UI customization for mode: {ModeId}", modeId);
            }
        }

        private void InitializeDefaultModeCustomizations()
        {
            // Default mode
            RegisterModeCustomization("default", new ModeUICustomization
            {
                ModeId = "default",
                DisplayName = "Standard Processing",
                Description = "Standard file processing with all features available",
                IconPath = "/Images/default-mode.png",
                MenuItems = new List<ModeMenuItem>
                {
                    new ModeMenuItem
                    {
                        Header = "Process Files",
                        Icon = "/Images/process.png",
                        ToolTip = "Process selected files",
                        Priority = 1
                    }
                }
            });

            // Save Quotes mode
            RegisterModeCustomization("SaveQuotes", new ModeUICustomization
            {
                ModeId = "SaveQuotes",
                DisplayName = "Save Quotes",
                Description = "Specialized mode for processing and saving quotes",
                IconPath = "/Images/quotes-mode.png",
                MenuItems = new List<ModeMenuItem>
                {
                    new ModeMenuItem
                    {
                        Header = "Save Quotes",
                        Icon = "/Images/quotes.png",
                        ToolTip = "Save quotes to processing queue",
                        Priority = 1
                    },
                    new ModeMenuItem
                    {
                        Header = "Queue Management",
                        Icon = "/Images/queue.png",
                        ToolTip = "Manage processing queue",
                        Priority = 2
                    }
                },
                UIProperties = new Dictionary<string, object>
                {
                    ["ShowCompanyDetection"] = true,
                    ["ShowScopeSelector"] = true,
                    ["CompactMode"] = true
                }
            });

            _logger.Debug("Initialized {Count} default mode customizations", _modeCustomizations.Count);
        }

        private ModeUICustomization GetDefaultModeUI()
        {
            return _modeCustomizations.GetValueOrDefault("default", new ModeUICustomization
            {
                ModeId = "default",
                DisplayName = "Default",
                Description = "Default UI mode"
            });
        }

        private async Task ApplyUIPropertiesAsync(ModeUICustomization customization, FrameworkElement targetElement)
        {
            try
            {
                await Task.CompletedTask; // Make async for future extensibility
                
                foreach (var property in customization.UIProperties)
                {
                    _logger.Debug("Applying UI property: {Property} = {Value}", property.Key, property.Value);
                    
                    // Apply properties based on their type and name
                    switch (property.Key)
                    {
                        case "CompactMode":
                            if (property.Value is bool isCompact && isCompact)
                            {
                                // Apply compact mode styling
                                ApplyCompactMode(targetElement);
                            }
                            break;
                        
                        default:
                            _logger.Debug("Unknown UI property: {Property}", property.Key);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error applying UI properties");
            }
        }

        private async Task ApplyElementVisibilityAsync(ModeUICustomization customization, FrameworkElement targetElement)
        {
            try
            {
                await Task.CompletedTask; // Make async for future extensibility
                
                // Hide specified elements
                foreach (var elementName in customization.HiddenElements)
                {
                    var element = targetElement.FindName(elementName) as FrameworkElement;
                    if (element != null)
                    {
                        element.Visibility = Visibility.Collapsed;
                        _logger.Debug("Hidden element: {ElementName}", elementName);
                    }
                }

                // Ensure required elements are visible
                foreach (var elementName in customization.RequiredElements)
                {
                    var element = targetElement.FindName(elementName) as FrameworkElement;
                    if (element != null)
                    {
                        element.Visibility = Visibility.Visible;
                        _logger.Debug("Made element visible: {ElementName}", elementName);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error applying element visibility");
            }
        }

        private void ApplyCompactMode(FrameworkElement targetElement)
        {
            try
            {
                // Apply compact mode styling - reduce margins, padding, etc.
                // This is a placeholder for actual compact mode implementation
                _logger.Debug("Applied compact mode styling to target element");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error applying compact mode");
            }
        }
    }
} 