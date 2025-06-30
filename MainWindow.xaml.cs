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
        /// Handle double-click on recent portion items to select them
        /// </summary>
        private void RecentPortionItem_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is ListBoxItem item && item.DataContext is string portion)
            {
                ViewModel.SelectedFileNamePortion = portion;
                ViewModel.FileNamePortionSearchText = portion;
                ViewModel.SelectFileNamePortionCommand.Execute(portion);
            }
        }
        
        /// <summary>
        /// Handle selection of filename portion from main list
        /// </summary>
        private void FileNamePortionItem_Selected(object sender, RoutedEventArgs e)
        {
            if (sender is ListBoxItem item && item.DataContext is string portion)
            {
                ViewModel.SelectFileNamePortionCommand.Execute(portion);
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