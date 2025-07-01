using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using DocHandler.Helpers;
using DocHandler.ViewModels;
using Serilog;
using DragEventArgs = System.Windows.DragEventArgs;

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
            
            // Clean up old Outlook temp files on startup
            OutlookAttachmentHelper.CleanupTempFiles();
        }
        
        private void Border_Drop(object sender, DragEventArgs e)
        {
            try
            {
                // Reset border appearance
                DropBorder.BorderBrush = (Brush)FindResource("SystemControlForegroundBaseMediumBrush");
                DropBorder.BorderThickness = new Thickness(2);
                
                var filesToAdd = new List<string>();
                
                // Check for standard file drop first
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    filesToAdd.AddRange(files);
                    _logger.Information("Received {Count} files via standard file drop", files.Length);
                }
                
                // Check for Outlook attachments
                if (e.Data.GetDataPresent("FileGroupDescriptor") || e.Data.GetDataPresent("FileGroupDescriptorW"))
                {
                    _logger.Information("Detected Outlook attachment drop");
                    
                    try
                    {
                        // Show processing indicator
                        Mouse.OverrideCursor = Cursors.Wait;
                        ViewModel.StatusMessage = "Extracting Outlook attachments...";
                        
                        // Extract Outlook attachments
                        var outlookFiles = OutlookAttachmentHelper.ExtractOutlookAttachments(e.Data);
                        
                        if (outlookFiles.Any())
                        {
                            filesToAdd.AddRange(outlookFiles);
                            _logger.Information("Successfully extracted {Count} Outlook attachments", outlookFiles.Count);
                            
                            // Mark these files for cleanup after processing
                            ViewModel.AddTempFilesForCleanup(outlookFiles);
                        }
                        else
                        {
                            _logger.Warning("No attachments could be extracted from Outlook drop");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to extract Outlook attachments");
                        MessageBox.Show(
                            "Failed to extract attachments from Outlook. Please try saving the attachments first.",
                            "Outlook Attachment Error",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                    finally
                    {
                        Mouse.OverrideCursor = null;
                    }
                }
                
                // Add all collected files
                if (filesToAdd.Any())
                {
                    ViewModel.AddFiles(filesToAdd.ToArray());
                }
                else if (!e.Data.GetDataPresent(DataFormats.FileDrop) && 
                         !e.Data.GetDataPresent("FileGroupDescriptor") && 
                         !e.Data.GetDataPresent("FileGroupDescriptorW"))
                {
                    // Log available formats for debugging
                    var formats = e.Data.GetFormats();
                    _logger.Debug("Available drop formats: {Formats}", string.Join(", ", formats));
                    
                    MessageBox.Show(
                        "The dropped items are not in a supported format. Please drop files or Outlook attachments.",
                        "Unsupported Format",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
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
            // Check if the drag data contains files or Outlook attachments
            if (e.Data.GetDataPresent(DataFormats.FileDrop) || 
                e.Data.GetDataPresent("FileGroupDescriptor") || 
                e.Data.GetDataPresent("FileGroupDescriptorW"))
            {
                e.Effects = DragDropEffects.Copy;
                
                // Highlight the border
                DropBorder.BorderBrush = (Brush)FindResource("SystemControlHighlightAccentBrush");
                DropBorder.BorderThickness = new Thickness(3);
                
                // Update status for Outlook attachments
                if (e.Data.GetDataPresent("FileGroupDescriptor") || e.Data.GetDataPresent("FileGroupDescriptorW"))
                {
                    ViewModel.StatusMessage = "Drop Outlook attachments here...";
                }
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
            
            // Reset status message
            ViewModel.UpdateUI();
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
        
        /// <summary>
        /// Handle saving preferences when menu item is clicked
        /// </summary>
        private void MenuItem_SavePreferences(object sender, RoutedEventArgs e)
        {
            ViewModel.SavePreferences();
        }
        
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Save window position
            ViewModel.SaveWindowState(Left, Top, Width, Height, WindowState.ToString());
            
            // Save preferences
            ViewModel.SavePreferences();
            
            // Cleanup
            ViewModel.Cleanup();
            
            // Clean up any remaining Outlook temp files
            OutlookAttachmentHelper.CleanupTempFiles();
        }
    }
}