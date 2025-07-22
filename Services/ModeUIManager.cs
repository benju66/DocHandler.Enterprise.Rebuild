using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Manages UI switching between different processing modes
    /// </summary>
    public class ModeUIManager : IAdvancedModeUIManager
    {
        private readonly ILogger _logger;
        private readonly IAdvancedModeUIProvider _modeUIProvider;
        private readonly IDynamicMenuBuilder _menuBuilder;
        private string? _currentMode;
        private FrameworkElement? _targetElement;

        public string? CurrentMode => _currentMode;

        public event EventHandler<AdvancedModeChangedEventArgs>? ModeChanged;

        public ModeUIManager(
            IAdvancedModeUIProvider modeUIProvider,
            IDynamicMenuBuilder menuBuilder)
        {
            _logger = Log.ForContext<ModeUIManager>();
            _modeUIProvider = modeUIProvider ?? throw new ArgumentNullException(nameof(modeUIProvider));
            _menuBuilder = menuBuilder ?? throw new ArgumentNullException(nameof(menuBuilder));
            
            _currentMode = "default";
            _logger.Debug("ModeUIManager initialized with default mode");
        }

        public async Task SwitchToModeAsync(string modeId)
        {
            try
            {
                _logger.Information("Switching to mode: {ModeId} from current mode: {CurrentMode}", modeId, _currentMode);

                if (string.IsNullOrEmpty(modeId))
                {
                    _logger.Warning("Mode ID is null or empty, switching to default mode");
                    modeId = "default";
                }

                var previousMode = _currentMode;
                
                // Check if mode is supported
                if (!_modeUIProvider.SupportsModeUI(modeId))
                {
                    _logger.Warning("Mode {ModeId} is not supported, falling back to default", modeId);
                    modeId = "default";
                }

                // Get the UI customization for the new mode
                var customization = await _modeUIProvider.GetModeUIAsync(modeId);
                
                // Apply UI changes if target element is set
                if (_targetElement != null)
                {
                    await _modeUIProvider.ApplyModeUIAsync(modeId, _targetElement);
                }

                // Update current mode
                _currentMode = modeId;

                // Fire mode changed event
                OnModeChanged(new AdvancedModeChangedEventArgs
                {
                    PreviousMode = previousMode,
                    CurrentMode = _currentMode,
                    UICustomization = customization
                });

                _logger.Information("Successfully switched to mode: {ModeId}", modeId);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error switching to mode: {ModeId}", modeId);
                
                // Try to fall back to default mode
                if (modeId != "default")
                {
                    _logger.Information("Attempting to fall back to default mode");
                    await SwitchToDefaultModeAsync();
                }
            }
        }

        public async Task SwitchToDefaultModeAsync()
        {
            try
            {
                _logger.Information("Switching to default mode");
                await SwitchToModeAsync("default");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error switching to default mode");
            }
        }

        public async Task<IEnumerable<string>> GetAvailableModesAsync()
        {
            try
            {
                var modes = await _modeUIProvider.GetAvailableModesAsync();
                _logger.Debug("Retrieved {Count} available modes", modes.Count());
                return modes;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error getting available modes");
                return new[] { "default" }; // Always return at least default mode
            }
        }

        /// <summary>
        /// Set the target UI element for mode switching
        /// </summary>
        public void SetTargetElement(FrameworkElement targetElement)
        {
            _targetElement = targetElement;
            _logger.Debug("Target element set for mode UI management");
        }

        /// <summary>
        /// Initialize the UI manager with the current mode
        /// </summary>
        public async Task InitializeAsync(string? initialMode = null)
        {
            try
            {
                var modeToSet = initialMode ?? "default";
                _logger.Information("Initializing ModeUIManager with mode: {ModeId}", modeToSet);
                
                await SwitchToModeAsync(modeToSet);
                
                _logger.Information("ModeUIManager initialized successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error initializing ModeUIManager");
            }
        }

        /// <summary>
        /// Get the current mode's UI customization
        /// </summary>
        public async Task<ModeUICustomization?> GetCurrentModeUIAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_currentMode))
                    return null;

                return await _modeUIProvider.GetModeUIAsync(_currentMode);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error getting current mode UI customization");
                return null;
            }
        }

        /// <summary>
        /// Build menu items for the current mode
        /// </summary>
        public async Task<IEnumerable<System.Windows.Controls.MenuItem>> GetCurrentModeMenuItemsAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_currentMode))
                    return Enumerable.Empty<System.Windows.Controls.MenuItem>();

                return await _menuBuilder.BuildMenuItemsAsync(_currentMode);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error building menu items for current mode");
                return Enumerable.Empty<System.Windows.Controls.MenuItem>();
            }
        }

        /// <summary>
        /// Build toolbar items for the current mode
        /// </summary>
        public async Task<IEnumerable<System.Windows.Controls.Control>> GetCurrentModeToolbarItemsAsync()
        {
            try
            {
                if (string.IsNullOrEmpty(_currentMode))
                    return Enumerable.Empty<System.Windows.Controls.Control>();

                return await _menuBuilder.BuildToolbarItemsAsync(_currentMode);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error building toolbar items for current mode");
                return Enumerable.Empty<System.Windows.Controls.Control>();
            }
        }

        /// <summary>
        /// Check if a specific mode is currently active
        /// </summary>
        public bool IsCurrentMode(string modeId)
        {
            return string.Equals(_currentMode, modeId, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Reset UI to default state
        /// </summary>
        public async Task ResetUIAsync()
        {
            try
            {
                _logger.Information("Resetting UI to default state");
                
                if (_targetElement != null)
                {
                    await _modeUIProvider.ResetUIAsync(_targetElement);
                }
                
                await SwitchToDefaultModeAsync();
                
                _logger.Information("UI reset completed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error resetting UI");
            }
        }

        protected virtual void OnModeChanged(AdvancedModeChangedEventArgs e)
        {
            try
            {
                ModeChanged?.Invoke(this, e);
                _logger.Debug("Mode changed event fired: {PreviousMode} -> {CurrentMode}", e.PreviousMode, e.CurrentMode);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error firing mode changed event");
            }
        }

        /// <summary>
        /// Dispose of resources
        /// </summary>
        public void Dispose()
        {
            try
            {
                _targetElement = null;
                ModeChanged = null;
                _logger.Debug("ModeUIManager disposed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error disposing ModeUIManager");
            }
        }
    }
} 