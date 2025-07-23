using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Centralized UI state management service extracted from MainViewModel.
    /// Handles UI thread synchronization, progress reporting, status updates, and complex UI state coordination.
    /// </summary>
    public class UIStateService : IUIStateService
    {
        private readonly ILogger _logger;
        private double _currentProgress = 0.0;
        private string _currentStatus = "Ready";
        private string _currentQueueStatus = "Drop quote documents";
        private bool _isProcessing = false;

        public UIStateService()
        {
            _logger = Log.ForContext<UIStateService>();
            _logger.Information("UIStateService initialized");
        }

        #region Progress Management

        public async Task UpdateProgressAsync(double progressValue, string? statusMessage = null)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _currentProgress = Math.Max(0, Math.Min(100, progressValue));
                
                if (!string.IsNullOrEmpty(statusMessage))
                {
                    _currentStatus = statusMessage;
                }

                _logger.Debug("Progress updated: {Progress}% - {Status}", 
                    _currentProgress, statusMessage ?? "No status change");
            });
        }

        public async Task ResetProgressAsync()
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _currentProgress = 0.0;
                _logger.Debug("Progress reset to 0%");
            });
        }

        public async Task<double> GetCurrentProgressAsync()
        {
            return await InvokeOnUIThreadAsync(() => _currentProgress);
        }

        #endregion Progress Management

        #region Status Management

        public async Task UpdateStatusAsync(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                _logger.Warning("Empty status message provided");
                return;
            }

            await InvokeOnUIThreadAsync(() =>
            {
                _currentStatus = message;
                _logger.Debug("Status updated: {Status}", message);
            });
        }

        public async Task UpdateQueueStatusAsync(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                _logger.Warning("Empty queue status message provided");
                return;
            }

            await InvokeOnUIThreadAsync(() =>
            {
                _currentQueueStatus = message;
                _logger.Debug("Queue status updated: {QueueStatus}", message);
            });
        }

        public async Task<string> GetCurrentStatusAsync()
        {
            return await InvokeOnUIThreadAsync(() => _currentStatus);
        }

        #endregion Status Management

        #region Processing State Management

        public async Task SetProcessingAsync(bool isProcessing)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _isProcessing = isProcessing;
                _logger.Debug("Processing state changed: {IsProcessing}", isProcessing);
            });
        }

        public async Task<bool> IsProcessingAsync()
        {
            return await InvokeOnUIThreadAsync(() => _isProcessing);
        }

        #endregion Processing State Management

        #region UI Synchronization

        public async Task InvokeOnUIThreadAsync(Action action)
        {
            if (action == null)
            {
                _logger.Warning("Null action provided to InvokeOnUIThreadAsync");
                return;
            }

            try
            {
                if (Application.Current?.Dispatcher?.CheckAccess() == true)
                {
                    // Already on UI thread
                    action();
                }
                else
                {
                    // Marshal to UI thread
                    await Application.Current.Dispatcher.InvokeAsync(action);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error executing action on UI thread");
                throw;
            }
        }

        public async Task<T> InvokeOnUIThreadAsync<T>(Func<T> function)
        {
            if (function == null)
            {
                _logger.Warning("Null function provided to InvokeOnUIThreadAsync<T>");
                throw new ArgumentNullException(nameof(function));
            }

            try
            {
                if (Application.Current?.Dispatcher?.CheckAccess() == true)
                {
                    // Already on UI thread
                    return function();
                }
                else
                {
                    // Marshal to UI thread
                    return await Application.Current.Dispatcher.InvokeAsync(function);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error executing function on UI thread");
                throw;
            }
        }

        #endregion UI Synchronization

        #region Complex UI State Updates

        public async Task RefreshUIStateAsync(UIStateContext context)
        {
            if (context == null)
            {
                _logger.Warning("Null UIStateContext provided");
                return;
            }

            await InvokeOnUIThreadAsync(() =>
            {
                try
                {
                    // This logic is extracted from MainViewModel.UpdateUI()
                    string statusMessage;
                    string buttonText;
                    bool canProcess;

                    if (context.SaveQuotesMode)
                    {
                        // Save Quotes Mode Logic
                        var hasCompanyName = !string.IsNullOrWhiteSpace(context.CompanyNameInput) || 
                                           !string.IsNullOrWhiteSpace(context.DetectedCompanyName);

                        canProcess = context.PendingFileCount > 0 && 
                                   context.AllFilesValid &&
                                   !context.IsProcessing && 
                                   !string.IsNullOrEmpty(context.SelectedScope) && 
                                   hasCompanyName;

                        buttonText = context.PendingFileCount > 1 ? "Process All Quotes" : "Process Quote";

                        if (context.PendingFileCount == 0)
                        {
                            statusMessage = "Save Quotes Mode: Drop quote documents";
                        }
                        else
                        {
                            statusMessage = $"{context.PendingFileCount} quote(s) ready to process";
                        }
                    }
                    else
                    {
                        // Standard Mode Logic
                        canProcess = context.PendingFileCount > 0 && context.AllFilesValid && !context.IsProcessing;
                        buttonText = context.PendingFileCount > 1 ? "Merge and Save" : "Process Files";

                        if (context.PendingFileCount == 0)
                        {
                            statusMessage = "Drop files here to begin";
                        }
                        else
                        {
                            statusMessage = $"{context.PendingFileCount} file(s) ready to process";
                        }
                    }

                    // Update internal state
                    _currentStatus = statusMessage;

                    _logger.Debug("UI state refreshed - CanProcess: {CanProcess}, ButtonText: {ButtonText}, Status: {Status}", 
                        canProcess, buttonText, statusMessage);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error refreshing UI state");
                    throw;
                }
            });
        }

        public async Task UpdateCanProcessStateAsync(bool canProcess, string? buttonText = null)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _logger.Debug("CanProcess state updated: {CanProcess}, ButtonText: {ButtonText}", 
                    canProcess, buttonText ?? "Not specified");
            });
        }

        #endregion Complex UI State Updates

        #region Error Display

        public async Task ShowErrorAsync(string title, string message)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _logger.Error("UI Error - {Title}: {Message}", title, message);
                
                try
                {
                    MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to show error dialog");
                }
            });
        }

        public async Task ShowWarningAsync(string title, string message)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _logger.Warning("UI Warning - {Title}: {Message}", title, message);
                
                try
                {
                    MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to show warning dialog");
                }
            });
        }

        public async Task ShowInfoAsync(string title, string message)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _logger.Information("UI Info - {Title}: {Message}", title, message);
                
                try
                {
                    MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to show info dialog");
                }
            });
        }

        #endregion Error Display

        #region Queue UI Management

        public async Task UpdateQueueCountAsync(int count)
        {
            await InvokeOnUIThreadAsync(() =>
            {
                var message = count > 0 ? $"{count} item(s) in queue" : "Drop quote documents";
                _currentQueueStatus = message;
                _logger.Debug("Queue count updated: {Count}", count);
            });
        }

        public async Task RefreshQueueUIAsync()
        {
            await InvokeOnUIThreadAsync(() =>
            {
                _logger.Debug("Queue UI refreshed");
                // Queue-specific refresh logic can be added here
            });
        }

        #endregion Queue UI Management
    }
} 