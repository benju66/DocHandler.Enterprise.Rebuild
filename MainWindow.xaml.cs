using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using DocHandler.Helpers;
using DocHandler.ViewModels;
using Serilog;
using DragEventArgs = System.Windows.DragEventArgs;
using System.Threading.Tasks;
using System.Diagnostics;

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
                // Validate stored position before applying
                var left = config.WindowLeft;
                var top = config.WindowTop;
                var width = config.WindowWidth;
                var height = config.WindowHeight;
                
                // Ensure window is at least partially visible on current screen configuration
                var virtualScreenLeft = SystemParameters.VirtualScreenLeft;
                var virtualScreenTop = SystemParameters.VirtualScreenTop;
                var virtualScreenWidth = SystemParameters.VirtualScreenWidth;
                var virtualScreenHeight = SystemParameters.VirtualScreenHeight;
                
                // Check if window would be visible
                var isVisible = left + width > virtualScreenLeft + 50 && // At least 50 pixels visible horizontally
                               left < virtualScreenLeft + virtualScreenWidth - 50 &&
                               top + height > virtualScreenTop + 50 && // At least 50 pixels visible vertically
                               top < virtualScreenTop + virtualScreenHeight - 50;
                
                if (isVisible)
                {
                    Left = left;
                    Top = top;
                    Width = width;
                    Height = height;
                    
                    // Only restore state if not minimized
                    if (Enum.TryParse<WindowState>(config.WindowState, out var state) && state != WindowState.Minimized)
                    {
                        WindowState = state;
                    }
                }
                else
                {
                    // Center window if saved position is not visible
                    WindowStartupLocation = WindowStartupLocation.CenterScreen;
                    _logger.Warning("Saved window position was off-screen, centering window");
                }
            }
            
            // Clean up old Outlook temp files on startup
            OutlookAttachmentHelper.CleanupTempFiles();
        }
        
        private void Border_Drop(object sender, DragEventArgs e)
        {
            // Cancel any running success animation immediately
            if (ViewModel.IsShowingSuccessAnimation)
            {
                CancelSuccessAnimation();
            }
            
            try
            {
                // Reset border appearance
                var dropBorder = sender as Border;
                if (dropBorder != null)
                {
                    dropBorder.BorderBrush = (Brush)FindResource("SystemControlForegroundBaseMediumBrush");
                    dropBorder.BorderThickness = new Thickness(2);
                }
                
                // DIAGNOSTIC: Log all available data formats
                var formats = e.Data.GetFormats();
                _logger.Information("Available drag formats: {Formats}", string.Join(", ", formats));
                
                var filesToAdd = new List<string>();
                
                // Check for standard file drop first - but handle COM exceptions
                try
                {
                    if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    {
                        string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                        if (files != null)
                        {
                            filesToAdd.AddRange(files);
                            _logger.Information("Received {Count} files via standard file drop", files.Length);
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    _logger.Warning("COM exception accessing FileDrop format: {Message}", comEx.Message);
                    // Continue trying other methods
                }
                
                // If no files yet, check for classic Outlook attachments
                if (!filesToAdd.Any() && (e.Data.GetDataPresent("FileGroupDescriptor") || e.Data.GetDataPresent("FileGroupDescriptorW")))
                {
                    _logger.Information("Detected classic Outlook attachment drop");
                    
                    try
                    {
                        Mouse.OverrideCursor = Cursors.Wait;
                        Dispatcher.InvokeAsync(() => ViewModel.StatusMessage = "Extracting Outlook attachments...");
                        
                        var outlookFiles = OutlookAttachmentHelper.ExtractOutlookAttachments(e.Data);
                        
                        if (outlookFiles.Any())
                        {
                            filesToAdd.AddRange(outlookFiles);
                            _logger.Information("Successfully extracted {Count} classic Outlook attachments", outlookFiles.Count);
                            ViewModel.AddTempFilesForCleanup(outlookFiles);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to extract classic Outlook attachments");
                    }
                    finally
                    {
                        Mouse.OverrideCursor = null;
                    }
                }
                
                // If still no files, try new Outlook formats
                if (!filesToAdd.Any())
                {
                    _logger.Information("Attempting to handle new Outlook formats");
                    
                    // Check for Chromium format (new Outlook indicator)
                    if (formats.Contains("Chromium Web Custom MIME Data Format"))
                    {
                        _logger.Information("Detected new Chromium-based Outlook");
                        
                        try
                        {
                            // Try to read the Chromium MIME data
                            var chromeData = e.Data.GetData("Chromium Web Custom MIME Data Format");
                            if (chromeData is MemoryStream ms)
                            {
                                var bytes = ms.ToArray();
                                _logger.Information("Chromium data size: {Size} bytes", bytes.Length);
                                
                                // FIRST: Log hex dump for analysis (do this before text conversions)
                                try
                                {
                                    var hexDump = BitConverter.ToString(bytes.Take(200).ToArray()).Replace("-", " ");
                                    _logger.Information("Hex dump (first 200 bytes): {Hex}", hexDump);
                                }
                                catch (Exception hexEx)
                                {
                                    _logger.Warning(hexEx, "Failed to create hex dump");
                                }
                                
                                // Try multiple parsing approaches
                                
                                // 1. Try as UTF-8 text
                                try
                                {
                                    var utf8Text = System.Text.Encoding.UTF8.GetString(bytes);
                                    // Replace non-printable characters to prevent logging issues
                                    var cleanUtf8 = System.Text.RegularExpressions.Regex.Replace(utf8Text, @"[\x00-\x1F\x7F-\x9F]", "?");
                                    _logger.Information("UTF-8 text (cleaned): {Text}", cleanUtf8.Length > 200 ? cleanUtf8.Substring(0, 200) + "..." : cleanUtf8);
                                }
                                catch (Exception utf8Ex) 
                                {
                                    _logger.Warning(utf8Ex, "Failed to parse as UTF-8");
                                }
                                
                                // 2. Try as UTF-16 text (this is what new Outlook uses)
                                try
                                {
                                    var utf16Text = System.Text.Encoding.Unicode.GetString(bytes);
                                    // Replace non-printable characters to prevent logging issues
                                    var cleanUtf16 = System.Text.RegularExpressions.Regex.Replace(utf16Text, @"[\x00-\x1F\x7F-\x9F]", "?");
                                    _logger.Information("UTF-16 text (cleaned): {Text}", cleanUtf16.Length > 200 ? cleanUtf16.Substring(0, 200) + "..." : cleanUtf16);
                                    
                                    // Try to parse as JSON to extract attachment info
                                    try
                                    {
                                        // Find the JSON start
                                        var jsonStart = utf16Text.IndexOf("{");
                                        if (jsonStart >= 0)
                                        {
                                            var jsonText = utf16Text.Substring(jsonStart);
                                            dynamic jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonText);
                                            
                                            if (jsonData?.attachmentFiles != null)
                                            {
                                                var attachmentNames = new List<string>();
                                                foreach (var file in jsonData.attachmentFiles)
                                                {
                                                    if (file.name != null)
                                                    {
                                                        attachmentNames.Add(file.name.ToString());
                                                    }
                                                }
                                                
                                                if (attachmentNames.Any())
                                                {
                                                    _logger.Information("Found attachments in new Outlook data: {Names}", string.Join(", ", attachmentNames));
                                                    
                                                    // Show user what files they're trying to drop
                                                    var fileList = string.Join("\n• ", attachmentNames);
                                                    MessageBox.Show(
                                                        $"The new Outlook is trying to share these attachments:\n\n• {fileList}\n\n" +
                                                        "Unfortunately, direct drag-and-drop from the new Outlook is not supported because the files are stored in the cloud.\n\n" +
                                                        "Please use one of these alternatives:\n" +
                                                        "• Save attachments to a folder first, then drag them here\n" +
                                                        "• Use the classic Outlook desktop application\n" +
                                                        "• Right-click attachments and select 'Save As'",
                                                        "New Outlook Attachments Detected",
                                                        MessageBoxButton.OK,
                                                        MessageBoxImage.Information);
                                                    return;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception jsonEx)
                                    {
                                        _logger.Warning(jsonEx, "Failed to parse Outlook JSON data");
                                    }
                                }
                                catch (Exception utf16Ex) 
                                {
                                    _logger.Warning(utf16Ex, "Failed to parse as UTF-16");
                                }
                                
                                // 3. Try as ASCII text
                                try
                                {
                                    var asciiText = System.Text.Encoding.ASCII.GetString(bytes);
                                    // Replace non-printable characters
                                    var cleanAscii = System.Text.RegularExpressions.Regex.Replace(asciiText, @"[\x00-\x1F\x7F-\xFF]", "?");
                                    _logger.Information("ASCII text (cleaned): {Text}", cleanAscii.Length > 200 ? cleanAscii.Substring(0, 200) + "..." : cleanAscii);
                                    
                                    // Check for MIME headers
                                    if (asciiText.Contains("Content-Type:") || asciiText.Contains("Content-Disposition:"))
                                    {
                                        _logger.Information("Found MIME headers in data");
                                    }
                                }
                                catch (Exception asciiEx)
                                {
                                    _logger.Warning(asciiEx, "Failed to parse as ASCII");
                                }
                                
                                // 4. Look for file paths in the binary data
                                try
                                {
                                    var possiblePaths = ExtractPathsFromBinary(bytes);
                                    if (possiblePaths.Any())
                                    {
                                        _logger.Information("Found possible file paths in Chromium data: {Paths}", string.Join(", ", possiblePaths));
                                        
                                        // Check if any of these paths are valid files
                                        foreach (var path in possiblePaths)
                                        {
                                            if (File.Exists(path))
                                            {
                                                filesToAdd.Add(path);
                                                _logger.Information("Found valid file: {Path}", path);
                                            }
                                        }
                                        
                                        if (filesToAdd.Any())
                                        {
                                            ViewModel.AddFiles(filesToAdd.ToArray());
                                            return;
                                        }
                                    }
                                    else
                                    {
                                        _logger.Information("No file paths found in binary data");
                                    }
                                }
                                catch (Exception pathEx)
                                {
                                    _logger.Warning(pathEx, "Failed to extract paths");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Failed to analyze Chromium format");
                        }
                        
                        // Show specific message for new Outlook
                        MessageBox.Show(
                            "Direct attachment drag-and-drop from the new Outlook is not currently supported.\n\n" +
                            "Please use one of these alternatives:\n" +
                            "• Save attachments to a folder first, then drag them here\n" +
                            "• Use the classic Outlook desktop application\n" +
                            "• Right-click attachments and select 'Save As'",
                            "New Outlook Not Supported",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information);
                        return;
                    }
                }
                
                // Add all collected files
                if (filesToAdd.Any())
                {
                    ViewModel.AddFiles(filesToAdd.ToArray());
                }
                else if (!formats.Contains("Chromium Web Custom MIME Data Format"))
                {
                    // Generic error for other unsupported formats
                    var availableFormats = string.Join(", ", formats);
                    _logger.Warning("No files could be extracted. Available formats: {Formats}", availableFormats);
                    
                    MessageBox.Show(
                        "The dropped items are not in a supported format.\n\n" +
                        "Please drop PDF, Word, or Excel files.",
                        "Unsupported Format",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error handling file drop");
                MessageBox.Show($"An error occurred while adding files:\n\n{ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private List<string> ExtractPathsFromBinary(byte[] bytes)
        {
            var paths = new List<string>();
            var text = System.Text.Encoding.ASCII.GetString(bytes);
            
            // Look for common path patterns
            var pathPatterns = new[]
            {
                @"[A-Za-z]:\\[^<>:""|?*\x00-\x1f]+\.[a-zA-Z]{2,4}",  // Windows paths
                @"\\\\[^<>:""|?*\x00-\x1f]+\.[a-zA-Z]{2,4}",         // UNC paths
                @"/[^<>:""|?*\x00-\x1f]+\.[a-zA-Z]{2,4}"             // Unix paths
            };
            
            foreach (var pattern in pathPatterns)
            {
                var matches = System.Text.RegularExpressions.Regex.Matches(text, pattern);
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    paths.Add(match.Value);
                }
            }
            
            // Also try to extract from UTF-16
            var utf16Text = System.Text.Encoding.Unicode.GetString(bytes);
            foreach (var pattern in pathPatterns)
            {
                var matches = System.Text.RegularExpressions.Regex.Matches(utf16Text, pattern);
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    paths.Add(match.Value);
                }
            }
            
            return paths.Distinct().ToList();
        }
        
        private void Border_DragEnter(object sender, DragEventArgs e)
        {
            var dropBorder = sender as Border;
            if (dropBorder == null) return;
            
            var formats = e.Data.GetFormats();
            
            // Check if it's the new Outlook (Chromium-based)
            if (formats.Contains("Chromium Web Custom MIME Data Format"))
            {
                e.Effects = DragDropEffects.None;
                
                // Still show visual feedback but with a different message
                dropBorder.BorderBrush = (Brush)FindResource("SystemControlForegroundBaseMediumHighBrush");
                dropBorder.BorderThickness = new Thickness(3);
                Dispatcher.InvokeAsync(() => ViewModel.StatusMessage = "New Outlook not supported - please save attachments first");
                return;
            }
            
            // Check if the drag data contains files or classic Outlook attachments
            if (e.Data.GetDataPresent(DataFormats.FileDrop) || 
                e.Data.GetDataPresent("FileGroupDescriptor") || 
                e.Data.GetDataPresent("FileGroupDescriptorW"))
            {
                e.Effects = DragDropEffects.Copy;
                
                // Highlight the border
                dropBorder.BorderBrush = (Brush)FindResource("SystemControlHighlightAccentBrush");
                dropBorder.BorderThickness = new Thickness(3);
                
                // Update status for Outlook attachments
                if (e.Data.GetDataPresent("FileGroupDescriptor") || e.Data.GetDataPresent("FileGroupDescriptorW"))
                {
                    Dispatcher.InvokeAsync(() => ViewModel.StatusMessage = "Drop Outlook attachments here...");
                }
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
        }
        
        private void Border_DragLeave(object sender, DragEventArgs e)
        {
            var dropBorder = sender as Border;
            if (dropBorder == null) return;
            
            // Reset border appearance
            dropBorder.BorderBrush = (Brush)FindResource("SystemControlForegroundBaseMediumBrush");
            dropBorder.BorderThickness = new Thickness(2);
            
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
                ViewModel.ScopeSearchText = "";  // Clear search when selecting from recent
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

        #region Save Quotes Success Animation

        private Storyboard? _currentSuccessAnimation;
        private DispatcherTimer? _animationResetTimer;

        /// <summary>
        /// Shows the success animation for Save Quotes mode
        /// </summary>
        public async Task ShowSaveQuotesSuccessAnimation(int savedCount = 1)
        {
            // Only show in Save Quotes mode
            if (!ViewModel.SaveQuotesMode) return;
            
            // Verify UI elements exist
            if (DropBorderSaveMode == null || SuccessOverlay == null || SuccessMessage == null)
            {
                _logger.Warning("Animation elements not found, skipping animation");
                return;
            }
            
            var stopwatch = Stopwatch.StartNew();
            
            // Cancel any existing animation
            CancelSuccessAnimation();
            
            try
            {
                // Get the storyboard with null check
                if (!(FindResource("SaveQuotesSuccessAnimation") is Storyboard successStoryboard))
                {
                    _logger.Warning("Animation storyboard not found");
                    return;
                }
                
                // Set animation state
                ViewModel.IsShowingSuccessAnimation = true;
                
                // Update success message
                SuccessMessage.Text = savedCount > 1 
                    ? $"Successfully saved {savedCount} quotes!" 
                    : "Successfully saved!";
                
                _currentSuccessAnimation = successStoryboard;
                
                // Show the success overlay
                await Dispatcher.InvokeAsync(() =>
                {
                    SuccessOverlay.Visibility = Visibility.Visible;
                    
                    // Initialize and fade in the checkmark and text using named StackPanel
                    SuccessStackPanel.Opacity = 0; // Set initial opacity programmatically
                    var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(500));
                    SuccessStackPanel.BeginAnimation(OpacityProperty, fadeIn);
                });
                
                // Start the border animation
                successStoryboard.Begin(DropBorderSaveMode);
                
                // Set up auto-reset timer (2.5 seconds total)
                _animationResetTimer = new DispatcherTimer
                {
                    Interval = TimeSpan.FromMilliseconds(2000) // Hold for 2 seconds after fade-in
                };
                _animationResetTimer.Tick += AnimationResetTimer_Tick;
                _animationResetTimer.Start();
                
                stopwatch.Stop();
                if (stopwatch.ElapsedMilliseconds > 50)
                {
                    _logger.Warning("Animation startup took {Duration}ms", 
                        stopwatch.ElapsedMilliseconds);
                }
            }
            catch (ResourceReferenceKeyNotFoundException ex)
            {
                _logger.Error(ex, "Animation resource not found");
                CancelSuccessAnimation();
            }
            catch (Exception ex)
            {
                // Log but don't crash on animation errors
                _logger.Error(ex, "Animation error");
                CancelSuccessAnimation();
            }
        }

        /// <summary>
        /// Cancels the current success animation immediately
        /// </summary>
        private void CancelSuccessAnimation()
        {
            // Stop any running animation
            _currentSuccessAnimation?.Stop();
            _currentSuccessAnimation = null;
            
            // Stop reset timer
            _animationResetTimer?.Stop();
            _animationResetTimer = null;
            
            // Reset visual state immediately
            Dispatcher.Invoke(() =>
            {
                // Hide overlay
                SuccessOverlay.Visibility = Visibility.Collapsed;
                // Reset StackPanel opacity using named reference
                SuccessStackPanel.Opacity = 0;
                
                // Reset border color instantly
                DropBorderSaveMode.BorderBrush = new SolidColorBrush(Color.FromRgb(0x19, 0x76, 0xD2));
                
                // Clear animation state
                ViewModel.IsShowingSuccessAnimation = false;
            });
        }

        /// <summary>
        /// Timer callback to reset animation after display time
        /// </summary>
        private void AnimationResetTimer_Tick(object? sender, EventArgs e)
        {
            _animationResetTimer?.Stop();
            
            // Fade out and reset
            Dispatcher.Invoke(async () =>
            {
                // Fade out the success overlay using named StackPanel
                var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(500));
                fadeOut.Completed += (s, args) =>
                {
                    SuccessOverlay.Visibility = Visibility.Collapsed;
                    ViewModel.IsShowingSuccessAnimation = false;
                };
                
                SuccessStackPanel.BeginAnimation(OpacityProperty, fadeOut);
                
                // Animate border back to normal
                var resetStoryboard = FindResource("SaveQuotesResetAnimation") as Storyboard;
                resetStoryboard?.Begin(DropBorderSaveMode);
            });
        }

        #endregion
        
        private void MainWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // Cancel any running animation
            CancelSuccessAnimation();
            
            // Save window position
            ViewModel.SaveWindowState(Left, Top, Width, Height, WindowState.ToString());
            
            // Save preferences
            ViewModel.SavePreferences();
            
            // Cleanup
            ViewModel.Cleanup();
            
            // Clean up any remaining Outlook temp files
            OutlookAttachmentHelper.CleanupTempFiles();
        }

        private void ScopeSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var vm = DataContext as ViewModels.MainViewModel;
            var textBox = sender as TextBox;
            
            if (vm == null || textBox == null) return;
            
            // If user starts typing and there's a selected scope, clear the selection
            // This allows them to search for a new scope
            if (!string.IsNullOrEmpty(vm.SelectedScope) && 
                textBox.Text != vm.SelectedScope && 
                !string.IsNullOrEmpty(textBox.Text))
            {
                vm.SelectedScope = null;
                // Update the ScopeSearchText to the current textbox content
                vm.ScopeSearchText = textBox.Text;
            }
            else if (string.IsNullOrEmpty(vm.SelectedScope))
            {
                // When no scope is selected, sync the search text
                vm.ScopeSearchText = textBox.Text;
            }
        }

        private void ScopeSearchTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            var vm = DataContext as ViewModels.MainViewModel;
            var textBox = sender as TextBox;
            
            if (vm == null || textBox == null) return;
            
            // If a scope is selected and user clicks the textbox, clear selection to enable search
            if (!string.IsNullOrEmpty(vm.SelectedScope))
            {
                vm.SelectedScope = null;
                vm.ScopeSearchText = "";
                // Focus and select all text for immediate typing
                textBox.Focus();
                textBox.SelectAll();
                e.Handled = true;
            }
        }

        private void ScopeSearchTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            var vm = DataContext as ViewModels.MainViewModel;
            var textBox = sender as TextBox;
            
            if (vm == null || textBox == null) return;
            
            // If a scope is selected and user types, clear selection and start fresh
            if (!string.IsNullOrEmpty(vm.SelectedScope))
            {
                vm.SelectedScope = null;
                vm.ScopeSearchText = e.Text; // Start with the typed character
                textBox.Clear();
                textBox.Text = e.Text;
                textBox.Select(e.Text.Length, 0); // Position cursor at end
                e.Handled = true;
            }
        }

        private void ScopeSearchTextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            var vm = DataContext as ViewModels.MainViewModel;
            var scopeSearchTextBox = sender as TextBox;
            var scopesListBox = GetCurrentScopesListBox();
            
            if (vm == null || scopesListBox == null) return;

            // Handle Ctrl+Enter for processing files
            if (e.Key == Key.Enter && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (vm.CanProcess)
                {
                    vm.ProcessFilesCommand.Execute(null);
                }
                e.Handled = true;
                return;
            }

            if (scopesListBox.Items.Count == 0) return;

            // Debug output
            System.Diagnostics.Debug.WriteLine($"PreviewKeyDown: {e.Key}, Items: {scopesListBox.Items.Count}");
            this.Title = $"DocHandler Enterprise - Key: {e.Key}";

            switch (e.Key)
            {
                case Key.Down:
                    NavigateDown(vm, scopesListBox);
                    e.Handled = true;
                    break;
                    
                case Key.Up:
                    NavigateUp(vm, scopesListBox);
                    e.Handled = true;
                    break;
                    
                case Key.Enter:
                    if (vm.SelectedScope != null)
                    {
                        vm.SelectScopeCommand.Execute(vm.SelectedScope);
                        // Don't clear the search box - let the UI binding show the selected scope
                        this.Title = "DocHandler Enterprise";
                    }
                    else if (vm.FilteredScopesOfWork.Count == 1)
                    {
                        // Auto-select the single filtered item
                        vm.SelectedScope = vm.FilteredScopesOfWork[0];
                        vm.SelectScopeCommand.Execute(vm.SelectedScope);
                        this.Title = "DocHandler Enterprise";
                    }
                    e.Handled = true;
                    break;
                    
                case Key.Escape:
                    // Clear both search and selection
                    vm.ScopeSearchText = "";
                    vm.SelectedScope = null;
                    this.Title = "DocHandler Enterprise";
                    e.Handled = true;
                    break;
                    
                case Key.Delete:
                case Key.Back:
                    // If user starts deleting with a scope selected, clear selection and position cursor
                    if (!string.IsNullOrEmpty(vm.SelectedScope))
                    {
                        vm.SelectedScope = null;
                        vm.ScopeSearchText = "";
                        // Position cursor at end for natural typing
                        scopeSearchTextBox?.Focus();
                        scopeSearchTextBox?.Select(0, 0); // Set cursor at beginning for immediate typing
                        e.Handled = true;
                    }
                    break;
            }
        }

        private ListBox GetCurrentScopesListBox()
        {
            // Try to find the save mode list box first
            var saveModeScopesListBox = FindName("ScopesListBoxSaveMode") as ListBox;
            if (saveModeScopesListBox != null && saveModeScopesListBox.IsVisible)
            {
                return saveModeScopesListBox;
            }
            
            // Fallback to normal mode (if it exists)
            return FindName("ScopesListBox") as ListBox;
        }

        private void NavigateDown(ViewModels.MainViewModel vm, ListBox scopesListBox)
        {
            var currentIndex = vm.FilteredScopesOfWork.IndexOf(vm.SelectedScope ?? "");
            var nextIndex = currentIndex < vm.FilteredScopesOfWork.Count - 1 ? currentIndex + 1 : 0;
            
            if (vm.FilteredScopesOfWork.Count > nextIndex)
            {
                vm.SelectedScope = vm.FilteredScopesOfWork[nextIndex];
                scopesListBox.ScrollIntoView(vm.SelectedScope);
                System.Diagnostics.Debug.WriteLine($"Down: Selected {vm.SelectedScope}");
            }
        }

        private void NavigateUp(ViewModels.MainViewModel vm, ListBox scopesListBox)
        {
            var currentIndex = vm.FilteredScopesOfWork.IndexOf(vm.SelectedScope ?? "");
            var prevIndex = currentIndex > 0 ? currentIndex - 1 : vm.FilteredScopesOfWork.Count - 1;
            
            if (vm.FilteredScopesOfWork.Count > prevIndex && prevIndex >= 0)
            {
                vm.SelectedScope = vm.FilteredScopesOfWork[prevIndex];
                scopesListBox.ScrollIntoView(vm.SelectedScope);
                System.Diagnostics.Debug.WriteLine($"Up: Selected {vm.SelectedScope}");
            }
        }
    }
}