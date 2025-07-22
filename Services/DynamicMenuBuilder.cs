using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Builds dynamic menus and toolbars based on mode configuration
    /// </summary>
    public class DynamicMenuBuilder : IDynamicMenuBuilder
    {
        private readonly ILogger _logger;
        private readonly IAdvancedModeUIProvider _modeUIProvider;
        private readonly Dictionary<string, Func<ModeMenuItem, MenuItem>> _menuFactories;
        private readonly Dictionary<string, Func<ModeMenuItem, Control>> _toolbarFactories;

        public DynamicMenuBuilder(IAdvancedModeUIProvider modeUIProvider)
        {
            _logger = Log.ForContext<DynamicMenuBuilder>();
            _modeUIProvider = modeUIProvider ?? throw new ArgumentNullException(nameof(modeUIProvider));
            _menuFactories = new Dictionary<string, Func<ModeMenuItem, MenuItem>>();
            _toolbarFactories = new Dictionary<string, Func<ModeMenuItem, Control>>();
            
            InitializeDefaultFactories();
            _logger.Debug("DynamicMenuBuilder initialized");
        }

        public async Task<IEnumerable<MenuItem>> BuildMenuItemsAsync(string modeId)
        {
            try
            {
                _logger.Debug("Building menu items for mode: {ModeId}", modeId);

                var customization = await _modeUIProvider.GetModeUIAsync(modeId);
                var menuItems = new List<MenuItem>();

                foreach (var modeMenuItem in customization.MenuItems.OrderBy(m => m.Priority))
                {
                    if (!modeMenuItem.IsVisible)
                        continue;

                    var menuItem = CreateMenuItem(modeMenuItem, modeId);
                    if (menuItem != null)
                    {
                        menuItems.Add(menuItem);
                    }
                }

                _logger.Debug("Built {Count} menu items for mode: {ModeId}", menuItems.Count, modeId);
                return menuItems;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error building menu items for mode: {ModeId}", modeId);
                return Enumerable.Empty<MenuItem>();
            }
        }

        public async Task<IEnumerable<Control>> BuildToolbarItemsAsync(string modeId)
        {
            try
            {
                _logger.Debug("Building toolbar items for mode: {ModeId}", modeId);

                var customization = await _modeUIProvider.GetModeUIAsync(modeId);
                var toolbarItems = new List<Control>();

                foreach (var modeMenuItem in customization.MenuItems.OrderBy(m => m.Priority))
                {
                    if (!modeMenuItem.IsVisible)
                        continue;

                    var toolbarItem = CreateToolbarItem(modeMenuItem, modeId);
                    if (toolbarItem != null)
                    {
                        toolbarItems.Add(toolbarItem);
                    }
                }

                _logger.Debug("Built {Count} toolbar items for mode: {ModeId}", toolbarItems.Count, modeId);
                return toolbarItems;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error building toolbar items for mode: {ModeId}", modeId);
                return Enumerable.Empty<Control>();
            }
        }

        public async Task<IEnumerable<MenuItem>> BuildContextMenuAsync(string modeId, object context)
        {
            try
            {
                _logger.Debug("Building context menu for mode: {ModeId} with context: {Context}", modeId, context?.GetType().Name);

                var customization = await _modeUIProvider.GetModeUIAsync(modeId);
                var contextMenuItems = new List<MenuItem>();

                // Filter menu items that are appropriate for context menus
                var contextItems = customization.MenuItems
                    .Where(m => m.IsVisible && IsContextAppropriate(m, context))
                    .OrderBy(m => m.Priority);

                foreach (var modeMenuItem in contextItems)
                {
                    var menuItem = CreateMenuItem(modeMenuItem, modeId);
                    if (menuItem != null)
                    {
                        contextMenuItems.Add(menuItem);
                    }
                }

                _logger.Debug("Built {Count} context menu items for mode: {ModeId}", contextMenuItems.Count, modeId);
                return contextMenuItems;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error building context menu for mode: {ModeId}", modeId);
                return Enumerable.Empty<MenuItem>();
            }
        }

        public void RegisterMenuItemFactory(string modeId, Func<ModeMenuItem, MenuItem> factory)
        {
            if (!string.IsNullOrEmpty(modeId) && factory != null)
            {
                _menuFactories[modeId] = factory;
                _logger.Debug("Registered menu item factory for mode: {ModeId}", modeId);
            }
        }

        public void RegisterToolbarItemFactory(string modeId, Func<ModeMenuItem, Control> factory)
        {
            if (!string.IsNullOrEmpty(modeId) && factory != null)
            {
                _toolbarFactories[modeId] = factory;
                _logger.Debug("Registered toolbar item factory for mode: {ModeId}", modeId);
            }
        }

        public void ClearFactories()
        {
            _menuFactories.Clear();
            _toolbarFactories.Clear();
            InitializeDefaultFactories();
            _logger.Debug("Cleared all factories and re-initialized defaults");
        }

        private void InitializeDefaultFactories()
        {
            // Default menu item factory
            _menuFactories["default"] = (modeMenuItem) => CreateDefaultMenuItem(modeMenuItem);
            
            // Default toolbar item factory  
            _toolbarFactories["default"] = (modeMenuItem) => CreateDefaultToolbarButton(modeMenuItem);
            
            _logger.Debug("Initialized default factories");
        }

        private MenuItem CreateMenuItem(ModeMenuItem modeMenuItem, string modeId)
        {
            try
            {
                // Try mode-specific factory first
                if (_menuFactories.TryGetValue(modeId, out var factory))
                {
                    return factory(modeMenuItem);
                }

                // Fall back to default factory
                if (_menuFactories.TryGetValue("default", out var defaultFactory))
                {
                    return defaultFactory(modeMenuItem);
                }

                // Last resort - create directly
                return CreateDefaultMenuItem(modeMenuItem);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error creating menu item: {Header}", modeMenuItem.Header);
                return null;
            }
        }

        private Control CreateToolbarItem(ModeMenuItem modeMenuItem, string modeId)
        {
            try
            {
                // Try mode-specific factory first
                if (_toolbarFactories.TryGetValue(modeId, out var factory))
                {
                    return factory(modeMenuItem);
                }

                // Fall back to default factory
                if (_toolbarFactories.TryGetValue("default", out var defaultFactory))
                {
                    return defaultFactory(modeMenuItem);
                }

                // Last resort - create directly
                return CreateDefaultToolbarButton(modeMenuItem);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error creating toolbar item: {Header}", modeMenuItem.Header);
                return null;
            }
        }

        private MenuItem CreateDefaultMenuItem(ModeMenuItem modeMenuItem)
        {
            var menuItem = new MenuItem
            {
                Header = modeMenuItem.Header,
                IsEnabled = modeMenuItem.IsEnabled,
                ToolTip = modeMenuItem.ToolTip
            };

            // Set icon if provided
            if (!string.IsNullOrEmpty(modeMenuItem.Icon))
            {
                try
                {
                    var image = new Image
                    {
                        Source = new BitmapImage(new Uri(modeMenuItem.Icon, UriKind.RelativeOrAbsolute)),
                        Width = 16,
                        Height = 16
                    };
                    menuItem.Icon = image;
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to load icon for menu item: {Icon}", modeMenuItem.Icon);
                }
            }

            // Set command if provided
            if (modeMenuItem.Command != null)
            {
                menuItem.Click += (sender, e) =>
                {
                    try
                    {
                        modeMenuItem.Command.Invoke(modeMenuItem.CommandParameter);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error executing menu item command: {Header}", modeMenuItem.Header);
                    }
                };
            }

            // Add sub-items if any
            foreach (var subItem in modeMenuItem.SubItems)
            {
                var subMenuItem = CreateDefaultMenuItem(subItem);
                if (subMenuItem != null)
                {
                    menuItem.Items.Add(subMenuItem);
                }
            }

            return menuItem;
        }

        private Control CreateDefaultToolbarButton(ModeMenuItem modeMenuItem)
        {
            var button = new Button
            {
                Content = modeMenuItem.Header,
                IsEnabled = modeMenuItem.IsEnabled,
                ToolTip = modeMenuItem.ToolTip,
                Margin = new Thickness(2, 2, 2, 2),
                Padding = new Thickness(8, 4, 8, 4),
                Style = Application.Current.FindResource("ToolbarButtonStyle") as Style
            };

            // Set icon if provided
            if (!string.IsNullOrEmpty(modeMenuItem.Icon))
            {
                try
                {
                    var stackPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal
                    };

                    var image = new Image
                    {
                        Source = new BitmapImage(new Uri(modeMenuItem.Icon, UriKind.RelativeOrAbsolute)),
                        Width = 16,
                        Height = 16,
                        Margin = new Thickness(0, 0, 4, 0)
                    };

                    var textBlock = new TextBlock
                    {
                        Text = modeMenuItem.Header,
                        VerticalAlignment = VerticalAlignment.Center
                    };

                    stackPanel.Children.Add(image);
                    stackPanel.Children.Add(textBlock);
                    button.Content = stackPanel;
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to load icon for toolbar button: {Icon}", modeMenuItem.Icon);
                }
            }

            // Set command if provided
            if (modeMenuItem.Command != null)
            {
                button.Click += (sender, e) =>
                {
                    try
                    {
                        modeMenuItem.Command.Invoke(modeMenuItem.CommandParameter);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error executing toolbar button command: {Header}", modeMenuItem.Header);
                    }
                };
            }

            return button;
        }

        private bool IsContextAppropriate(ModeMenuItem menuItem, object context)
        {
            // Determine if a menu item is appropriate for the given context
            // This can be extended with more sophisticated logic
            
            if (context == null)
                return true;

            // For now, all visible menu items are considered appropriate
            // This can be enhanced based on context type and menu item properties
            return menuItem.IsVisible;
        }
    }
} 