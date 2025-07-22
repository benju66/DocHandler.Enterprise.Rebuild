using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace DocHandler.Services
{
    /// <summary>
    /// Represents a UI menu item for mode-specific functionality
    /// </summary>
    public class ModeMenuItem
    {
        public string Header { get; set; } = string.Empty;
        public string Icon { get; set; } = string.Empty;
        public string ToolTip { get; set; } = string.Empty;
        public bool IsEnabled { get; set; } = true;
        public bool IsVisible { get; set; } = true;
        public int Priority { get; set; } = 0; // For ordering
        public Action<object>? Command { get; set; }
        public object? CommandParameter { get; set; }
        public List<ModeMenuItem> SubItems { get; set; } = new();
    }

    /// <summary>
    /// Represents UI customization options for a mode
    /// </summary>
    public class ModeUICustomization
    {
        public string ModeId { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string IconPath { get; set; } = string.Empty;
        public List<ModeMenuItem> MenuItems { get; set; } = new();
        public Dictionary<string, object> UIProperties { get; set; } = new();
        public List<string> HiddenElements { get; set; } = new();
        public List<string> RequiredElements { get; set; } = new();
    }

    /// <summary>
    /// Advanced UI customization provider for processing modes (Phase 2)
    /// </summary>
    public interface IAdvancedModeUIProvider
    {
        /// <summary>
        /// Get UI customization for a specific mode
        /// </summary>
        Task<ModeUICustomization> GetModeUIAsync(string modeId);

        /// <summary>
        /// Check if a mode supports UI customization
        /// </summary>
        bool SupportsModeUI(string modeId);

        /// <summary>
        /// Get all available modes with UI customization
        /// </summary>
        Task<IEnumerable<string>> GetAvailableModesAsync();

        /// <summary>
        /// Apply UI customization to the main window
        /// </summary>
        Task ApplyModeUIAsync(string modeId, FrameworkElement targetElement);

        /// <summary>
        /// Reset UI to default state
        /// </summary>
        Task ResetUIAsync(FrameworkElement targetElement);
    }

    /// <summary>
    /// Builds dynamic menus based on mode configuration
    /// </summary>
    public interface IDynamicMenuBuilder
    {
        /// <summary>
        /// Build menu items for a specific mode
        /// </summary>
        Task<IEnumerable<MenuItem>> BuildMenuItemsAsync(string modeId);

        /// <summary>
        /// Build toolbar items for a specific mode
        /// </summary>
        Task<IEnumerable<Control>> BuildToolbarItemsAsync(string modeId);

        /// <summary>
        /// Build context menu items for a specific mode
        /// </summary>
        Task<IEnumerable<MenuItem>> BuildContextMenuAsync(string modeId, object context);

        /// <summary>
        /// Register a menu item factory for a mode
        /// </summary>
        void RegisterMenuItemFactory(string modeId, Func<ModeMenuItem, MenuItem> factory);

        /// <summary>
        /// Clear all registered factories
        /// </summary>
        void ClearFactories();
    }

    /// <summary>
    /// Advanced UI manager for mode switching (Phase 2)
    /// </summary>
    public interface IAdvancedModeUIManager
    {
        /// <summary>
        /// Current active mode
        /// </summary>
        string? CurrentMode { get; }

        /// <summary>
        /// Switch to a specific mode
        /// </summary>
        Task SwitchToModeAsync(string modeId);

        /// <summary>
        /// Switch to default mode
        /// </summary>
        Task SwitchToDefaultModeAsync();

        /// <summary>
        /// Get available modes for UI switching
        /// </summary>
        Task<IEnumerable<string>> GetAvailableModesAsync();

        /// <summary>
        /// Initialize the UI manager with the current mode
        /// </summary>
        Task InitializeAsync(string? initialMode = null);

        /// <summary>
        /// Set the target UI element for mode switching
        /// </summary>
        void SetTargetElement(FrameworkElement targetElement);

        /// <summary>
        /// Event fired when mode changes
        /// </summary>
        event EventHandler<AdvancedModeChangedEventArgs> ModeChanged;
    }

    /// <summary>
    /// Event args for advanced mode change events
    /// </summary>
    public class AdvancedModeChangedEventArgs : EventArgs
    {
        public string? PreviousMode { get; set; }
        public string CurrentMode { get; set; } = string.Empty;
        public ModeUICustomization? UICustomization { get; set; }
    }
} 