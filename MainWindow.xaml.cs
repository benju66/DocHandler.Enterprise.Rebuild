using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using DocHandler.ViewModels;
using Serilog;

namespace DocHandler
{
    public partial class MainWindow : Window
    {
        private readonly ILogger _logger;
        private MainViewModel ViewModel => (MainViewModel)DataContext;
        
        public MainWindow()
        {
            InitializeComponent();
            _logger = Log.ForContext<MainWindow>();
            
            // Window closing event to cleanup
            Closing += MainWindow_Closing;
            
            // Restore window position after loading
            Loaded += MainWindow_Loaded;
        }
        
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Restore window position from config
            var config = ViewModel.ConfigService.Config;
            if (config.RememberWindowPosition)
            {
                Left = config.WindowLeft;
                Top = config.WindowTop;
                Width = config.WindowWidth;
                Height = config.WindowHeight;
                
                if (Enum.TryParse<WindowState>(config.WindowState, out var state))
                {
                    WindowState = state;
                }
            }
        }
        
        private void Border_Drop(object sender, DragEventArgs e)
        {
            try
            {
                // Reset border appearance
                DropBorder.BorderBrush = (Brush)FindResource("SystemControlForegroundBaseMediumBrush");
                DropBorder.BorderThickness = new Thickness(2);
                
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    ViewModel.AddFiles(files);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error handling file drop");
                MessageBox.Show("An error occurred while adding files.", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                
                // Highlight the border
                DropBorder.BorderBrush = (Brush)FindResource("SystemControlHighlightAccentBrush");
                DropBorder.BorderThickness = new Thickness(3);
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }
        
        private void Border_DragLeave(object sender, DragEventArgs e)
        {
            // Reset border appearance
            DropBorder.BorderBrush = (Brush)FindResource("SystemControlForegroundBaseMediumBrush");
            DropBorder.BorderThickness = new Thickness(2);
        }
        
        /// <summary>
        /// Handle double-click on recent scope items to select them
        /// </summary>
        private void RecentScopeItem_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListBoxItem item && item.DataContext is string scope)
            {
                ViewModel.SelectedScope = scope;
                ViewModel.ScopeSearchText = scope;
                ViewModel.SelectScopeCommand.Execute(scope);
            }
        }
        
        /// <summary>
        /// Handle selection of scope from main list
        /// </summary>
        private void ScopeItem_Selected(object sender, RoutedEventArgs e)
        {
            if (sender is ListBoxItem item && item.DataContext is string scope)
            {
                ViewModel.SelectScopeCommand.Execute(scope);
            }
        }
        
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Save window position
            ViewModel.SaveWindowState(Left, Top, Width, Height, WindowState.ToString());
            
            // Cleanup
            ViewModel.Cleanup();
        }
    }
}