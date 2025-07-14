using System;
using System.Collections.Concurrent;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using DocHandler.Models;
using Serilog;
using System.Collections.Generic; // Added missing import

namespace DocHandler.Services
{
    public partial class SaveQuotesQueueService : ObservableObject
    {
        private readonly ILogger _logger;
        private readonly ConcurrentQueue<SaveQuoteItem> _queue;
        private readonly SemaphoreSlim _processingSemaphore;
        private readonly ObservableCollection<SaveQuoteItem> _allItems;
        private readonly object _itemsLock = new object();
        private readonly OptimizedFileProcessingService _fileProcessingService;
        private CancellationTokenSource _cancellationTokenSource;
        
        public ObservableCollection<SaveQuoteItem> AllItems => _allItems;
        
        [ObservableProperty]
        private int _totalCount;
        
        [ObservableProperty]
        private int _processedCount;
        
        [ObservableProperty]
        private int _failedCount;
        
        [ObservableProperty]
        private bool _isProcessing;
        
        // Events
        public event EventHandler<SaveQuoteProgressEventArgs> ProgressChanged;
        public event EventHandler<SaveQuoteCompletedEventArgs> ItemCompleted;
        public event EventHandler QueueEmpty;
        public event EventHandler<string> StatusMessageChanged;
        
        public SaveQuotesQueueService()
        {
            _logger = Log.ForContext<SaveQuotesQueueService>();
            _queue = new ConcurrentQueue<SaveQuoteItem>();
            _processingSemaphore = new SemaphoreSlim(3, 3); // Max 3 concurrent operations
            _allItems = new ObservableCollection<SaveQuoteItem>();
            _fileProcessingService = new OptimizedFileProcessingService();
            _cancellationTokenSource = new CancellationTokenSource();
        }
        
        public void AddToQueue(FileItem file, string scope, string companyName, string saveLocation)
        {
            var item = new SaveQuoteItem
            {
                Id = Guid.NewGuid().ToString(),
                File = file,
                Scope = scope,
                CompanyName = companyName,
                SaveLocation = saveLocation,
                Status = SaveQuoteStatus.Queued,
                QueuedAt = DateTime.Now
            };
            
            _queue.Enqueue(item);
            
            Application.Current.Dispatcher.Invoke(() =>
            {
                lock (_itemsLock)
                {
                    _allItems.Add(item);
                    TotalCount = _allItems.Count;
                }
            });
            
            _logger.Information("Added item to queue: {File} - {Scope} - {Company}", 
                file.FileName, scope, companyName);
        }
        
        public async Task StartProcessingAsync()
        {
            if (IsProcessing) return;
            
            IsProcessing = true;
            _cancellationTokenSource = new CancellationTokenSource();
            
            StatusMessageChanged?.Invoke(this, "Processing queue...");
            
            try
            {
                await ProcessQueueAsync();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error processing queue");
            }
            finally
            {
                IsProcessing = false;
            }
        }
        
        private async Task ProcessQueueAsync()
        {
            const int maxConcurrency = 3;
            var semaphore = new SemaphoreSlim(maxConcurrency);
            var tasks = new List<Task>();
            
            while (!_cancellationTokenSource.Token.IsCancellationRequested)
            {
                if (_queue.TryDequeue(out var item))
                {
                    var task = ProcessItemAsync(item, semaphore);
                    tasks.Add(task);
                    
                    // Clean up completed tasks
                    tasks.RemoveAll(t => t.IsCompleted);
                    
                    // Don't let too many tasks accumulate
                    if (tasks.Count >= maxConcurrency * 2)
                    {
                        await Task.WhenAny(tasks);
                    }
                }
                else if (tasks.Count == 0)
                {
                    // Queue is empty and no tasks running
                    break;
                }
                else
                {
                    // Wait for any task to complete
                    await Task.WhenAny(tasks);
                }
            }
            
            // Wait for all remaining tasks
            await Task.WhenAll(tasks);
            
            // Fire queue empty event
            Application.Current.Dispatcher.Invoke(() =>
            {
                IsProcessing = false;
                QueueEmpty?.Invoke(this, EventArgs.Empty);
            });
        }
        
        private async Task ProcessItemAsync(SaveQuoteItem item, SemaphoreSlim semaphore)
        {
            await semaphore.WaitAsync();
            
            try
            {
                // Update status to processing
                Application.Current.Dispatcher.Invoke(() =>
                {
                    item.Status = SaveQuoteStatus.Processing;
                    UpdateCounts();
                });
                
                // Build output path
                var outputFileName = $"{item.Scope} - {item.CompanyName}.pdf";
                var outputPath = Path.Combine(item.SaveLocation, outputFileName);
                
                // Ensure unique filename
                outputPath = Path.Combine(item.SaveLocation, 
                    _fileProcessingService.GetUniqueFileName(item.SaveLocation, outputFileName));
                
                // Process the file
                var result = await _fileProcessingService.ConvertSingleFile(item.File.FilePath, outputPath);
                
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (result.Success)
                    {
                        item.Status = SaveQuoteStatus.Completed;
                        item.CompletedAt = DateTime.Now;
                        ProcessedCount++;
                        
                        ItemCompleted?.Invoke(this, new SaveQuoteCompletedEventArgs 
                        { 
                            Item = item, 
                            Success = true 
                        });
                    }
                    else
                    {
                        item.Status = SaveQuoteStatus.Failed;
                        item.ErrorMessage = result.ErrorMessage;
                        item.CompletedAt = DateTime.Now;
                        FailedCount++;
                        
                        ItemCompleted?.Invoke(this, new SaveQuoteCompletedEventArgs 
                        { 
                            Item = item, 
                            Success = false, 
                            ErrorMessage = result.ErrorMessage 
                        });
                    }
                    
                    UpdateCounts();
                });
                
                _logger.Information("Processed queue item: {File} - {Status}", 
                    item.File.FileName, item.Status);
            }
            catch (Exception ex)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    item.Status = SaveQuoteStatus.Failed;
                    item.ErrorMessage = ex.Message;
                    item.CompletedAt = DateTime.Now;
                    FailedCount++;
                    UpdateCounts();
                });
                
                _logger.Error(ex, "Failed to process queue item: {File}", item.File.FileName);
            }
            finally
            {
                semaphore.Release();
            }
        }
        
        private void UpdateCounts()
        {
            ProgressChanged?.Invoke(this, new SaveQuoteProgressEventArgs
            {
                TotalCount = TotalCount,
                ProcessedCount = ProcessedCount,
                FailedCount = FailedCount,
                IsProcessing = IsProcessing
            });
        }
        
        public void CancelItem(SaveQuoteItem item)
        {
            if (item.Status == SaveQuoteStatus.Queued)
            {
                Application.Current.Dispatcher.Invoke(() =>
                {
                    item.Status = SaveQuoteStatus.Cancelled;
                    UpdateCounts();
                });
                
                _logger.Information("Cancelled queue item: {File}", item.File.FileName);
            }
        }
        
        public void ClearCompleted()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                lock (_itemsLock)
                {
                    var completedItems = _allItems.Where(i => 
                        i.Status == SaveQuoteStatus.Completed || 
                        i.Status == SaveQuoteStatus.Cancelled).ToList();
                    
                    foreach (var item in completedItems)
                    {
                        _allItems.Remove(item);
                    }
                    
                    TotalCount = _allItems.Count;
                    ProcessedCount = _allItems.Count(i => i.Status == SaveQuoteStatus.Completed);
                    FailedCount = _allItems.Count(i => i.Status == SaveQuoteStatus.Failed);
                    
                    UpdateCounts();
                }
            });
            
            _logger.Information("Cleared completed queue items");
        }
    }
    
    public partial class SaveQuoteItem : ObservableObject
    {
        public string Id { get; set; } = "";
        public FileItem File { get; set; } = new();
        public string Scope { get; set; } = "";
        public string CompanyName { get; set; } = "";
        public string SaveLocation { get; set; } = "";
        
        [ObservableProperty]
        private SaveQuoteStatus _status;
        
        [ObservableProperty]
        private string? _errorMessage;
        
        public DateTime QueuedAt { get; set; }
        public DateTime? CompletedAt { get; set; }
    }
    
    public enum SaveQuoteStatus
    {
        Queued,
        Processing,
        Completed,
        Failed,
        Cancelled
    }
    
    public class SaveQuoteProgressEventArgs : EventArgs
    {
        public int TotalCount { get; set; }
        public int ProcessedCount { get; set; }
        public int FailedCount { get; set; }
        public bool IsProcessing { get; set; }
    }
    
    public class SaveQuoteCompletedEventArgs : EventArgs
    {
        public SaveQuoteItem Item { get; set; } = new();
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
    }
} 