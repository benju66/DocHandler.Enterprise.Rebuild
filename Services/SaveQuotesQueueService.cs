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
using DocHandler.Services;
using Serilog;
using System.Collections.Generic; // Added missing import

namespace DocHandler.Services
{
    public partial class SaveQuotesQueueService : ObservableObject, IDisposable
    {
        private readonly ILogger _logger;
        private readonly ConcurrentQueue<SaveQuoteItem> _queue;
        private readonly SemaphoreSlim _processingSemaphore;
        private readonly ObservableCollection<SaveQuoteItem> _allItems;
        private readonly object _itemsLock = new object();
        private readonly OptimizedFileProcessingService _fileProcessingService;
        private CancellationTokenSource _cancellationTokenSource;
        private readonly ConfigurationService _configService;
        private readonly PdfCacheService _pdfCacheService;
        private readonly ProcessManager _processManager;
        private readonly StaThreadPool _staThreadPool; // Added STA thread pool for COM operations
        
        // Add private disposal tracking
        private bool _disposed = false;
        
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
        public event EventHandler<SaveQuoteProgressEventArgs>? ProgressChanged;
        public event EventHandler<SaveQuoteCompletedEventArgs>? ItemCompleted;
        public event EventHandler? QueueEmpty;
        public event EventHandler<string>? StatusMessageChanged;
        
        /// <summary>
        /// Stop the queue processing and cleanup resources
        /// </summary>
        public void StopProcessing()
        {
            _logger.Information("Stopping queue processing");
            _cancellationTokenSource?.Cancel();
        }
        
        public SaveQuotesQueueService(ConfigurationService configService, PdfCacheService pdfCacheService, ProcessManager processManager, OptimizedFileProcessingService sharedFileProcessingService)
        {
            _logger = Log.ForContext<SaveQuotesQueueService>();
            _configService = configService;
            _pdfCacheService = pdfCacheService;
            _processManager = processManager;
            
            // Initialize the queue
            _queue = new ConcurrentQueue<SaveQuoteItem>();
            _allItems = new ObservableCollection<SaveQuoteItem>();
            
            // CRITICAL FIX: Initialize the CancellationTokenSource
            _cancellationTokenSource = new CancellationTokenSource();
            
            // Use shared file processing service instead of creating new instance
            _fileProcessingService = sharedFileProcessingService ?? throw new ArgumentNullException(nameof(sharedFileProcessingService));
            
            // Initialize STA thread pool for COM operations
            _staThreadPool = new StaThreadPool(1, "SaveQuotesQueue");
            
            // Determine optimal concurrency (conservative approach)
            var maxConcurrency = Math.Min(Environment.ProcessorCount - 1, 3);
            _processingSemaphore = new SemaphoreSlim(maxConcurrency, maxConcurrency);
            
            _logger.Information("Queue service initialized with max concurrency: {MaxConcurrency}", maxConcurrency);
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
            
            // THREADING FIX: Use BeginInvoke to avoid blocking caller thread
            Application.Current.Dispatcher.BeginInvoke(() =>
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
            _logger.Information("StartProcessingAsync called, IsProcessing: {IsProcessing}", IsProcessing);
            
            if (IsProcessing) 
            {
                _logger.Warning("Already processing, returning");
                return;
            }
            
            IsProcessing = true;
            
            // Ensure we have a fresh cancellation token
            _cancellationTokenSource?.Cancel();
            _cancellationTokenSource?.Dispose();
            _cancellationTokenSource = new CancellationTokenSource();
            
            StatusMessageChanged?.Invoke(this, "Processing queue...");
            
            try
            {
                _logger.Information("Starting ProcessQueueAsync");
                await ProcessQueueAsync();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error processing queue");
                throw; // Re-throw to surface the error
            }
            finally
            {
                IsProcessing = false;
                _logger.Information("Queue processing completed");
            }
        }
        
        private async Task ProcessQueueAsync()
        {
            var maxConcurrency = _configService.Config.MaxParallelProcessing;
            var semaphore = new SemaphoreSlim(maxConcurrency);
            var tasks = new List<Task>();
            
            _logger.Information("Processing queue with {MaxConcurrency} parallel tasks, queue has {Count} items", 
                maxConcurrency, _queue.Count);
            
            // Create a single converter for the entire batch
            using (var batchConverter = new ReliableOfficeConverter())
            {
                try
                {
                    while (!_cancellationTokenSource.Token.IsCancellationRequested)
                    {
                        if (_queue.TryDequeue(out var item))
                        {
                            _logger.Debug("Dequeued item: {FileName} for processing", item.File.FileName);
                            var task = ProcessItemWithConverterAsync(item, semaphore, batchConverter);
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
                            // Queue is empty and no tasks running - we're done!
                            _logger.Information("Queue is empty and all tasks completed");
                            break;
                        }
                        else
                        {
                            // Wait for any task to complete before checking queue again
                            await Task.WhenAny(tasks);
                            tasks.RemoveAll(t => t.IsCompleted);
                        }
                    }
                    
                    // Wait for all remaining tasks
                    if (tasks.Count > 0)
                    {
                        _logger.Information("Waiting for {Count} remaining tasks to complete", tasks.Count);
                        await Task.WhenAll(tasks);
                    }
                }
                finally
                {
                    // CRITICAL: Always finish batch to cleanup Office instances
                    _logger.Information("About to call FinishBatch on converter");
                    batchConverter.FinishBatch();
                    _logger.Information("FinishBatch called successfully - Office instances should be cleaned up");
                }
            }
            
            // REMOVED: OnQueueProcessingCompleted call since method no longer exists
            // ReliableOfficeConverter instances are already disposed after each batch
            
            // THREADING FIX: Use BeginInvoke for queue empty event to avoid blocking
            Application.Current.Dispatcher.BeginInvoke(() =>
            {
                IsProcessing = false;
                QueueEmpty?.Invoke(this, EventArgs.Empty);
            });
        }
        
        private async Task ProcessItemWithConverterAsync(SaveQuoteItem item, SemaphoreSlim semaphore, ReliableOfficeConverter converter)
        {
            await semaphore.WaitAsync();
            
            try
            {
                // THREADING FIX: Use BeginInvoke to avoid blocking the processing thread
                Application.Current.Dispatcher.BeginInvoke(() =>
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
                
                _logger.Information("=== STARTING QUEUE ITEM PROCESSING ===");
                _logger.Information("QUEUE: Processing file: {FileName} ({Extension})", 
                    item.File.FileName, Path.GetExtension(item.File.FilePath));
                _logger.Information("QUEUE: Output path: {OutputPath}", outputPath);

                // Execute conversion on STA thread
                _logger.Information("QUEUE: Starting conversion on STA thread (Current Thread: {ThreadId})", 
                    Thread.CurrentThread.ManagedThreadId);

                var result = await _staThreadPool.ExecuteAsync(() =>
                {
                    _logger.Information("QUEUE: Now executing on STA thread {ThreadId} (Apartment: {ApartmentState})", 
                        Thread.CurrentThread.ManagedThreadId, Thread.CurrentThread.GetApartmentState());
                    
                    // Process based on file type
                    var extension = Path.GetExtension(item.File.FilePath).ToLowerInvariant();
                    
                    if (extension == ".pdf")
                    {
                        // Just copy PDF files
                        File.Copy(item.File.FilePath, outputPath, true);
                        return new ConversionResult { Success = true, OutputPath = outputPath };
                    }
                    else if (extension == ".doc" || extension == ".docx")
                    {
                        // Use batch converter for Word files
                        return converter.ConvertWordToPdf(item.File.FilePath, outputPath, singleUse: false);
                    }
                    else if (extension == ".xls" || extension == ".xlsx")
                    {
                        // Use batch converter for Excel files
                        return converter.ConvertExcelToPdf(item.File.FilePath, outputPath, singleUse: false);
                    }
                    else
                    {
                        return new ConversionResult
                        {
                            Success = false,
                            ErrorMessage = $"Unsupported file type: {extension}"
                        };
                    }
                });

                _logger.Information("QUEUE: Conversion completed - Success: {Success}, Error: {Error}", 
                    result.Success, result.ErrorMessage ?? "None");

                _logger.Information("=== QUEUE ITEM PROCESSING COMPLETED ===");
                
                // THREADING FIX: Use BeginInvoke to avoid blocking the processing thread
                Application.Current.Dispatcher.BeginInvoke(() =>
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
                // THREADING FIX: Use BeginInvoke to avoid blocking the processing thread
                Application.Current.Dispatcher.BeginInvoke(() =>
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
        
        // Add method to update concurrency at runtime
        public void UpdateMaxConcurrency(int newMax)
        {
            if (newMax < 1) newMax = 1;
            if (newMax > 10) newMax = 10;
            
            _configService.Config.MaxParallelProcessing = newMax;
            _ = _configService.SaveConfiguration();
            
            _logger.Information("Updated max concurrency to {MaxConcurrency}", newMax);
            
            // Note: This will take effect on next queue processing
        }
        
        public void CancelItem(SaveQuoteItem item)
        {
            if (item.Status == SaveQuoteStatus.Queued)
            {
                // THREADING FIX: Use BeginInvoke to avoid blocking caller thread
                Application.Current.Dispatcher.BeginInvoke(() =>
                {
                    item.Status = SaveQuoteStatus.Cancelled;
                    UpdateCounts();
                });
                
                _logger.Information("Cancelled queue item: {File}", item.File.FileName);
            }
        }
        
        public void ClearCompleted()
        {
            // THREADING FIX: Use BeginInvoke to avoid blocking caller thread
            Application.Current.Dispatcher.BeginInvoke(() =>
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

        // Implement IDisposable
        public void Dispose()
        {
            _logger.Information("Disposing SaveQuotesQueueService");
            
            // Cancel any ongoing processing
            _cancellationTokenSource?.Cancel();
            _cancellationTokenSource?.Dispose();
            
            _staThreadPool?.Dispose();
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                try
                {
                    // Cancel any pending operations
                    _cancellationTokenSource?.Cancel();
                    _logger.Information("Queue processing cancelled");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error cancelling queue operations");
                }
                
                try
                {
                    _cancellationTokenSource?.Dispose();
                    _logger.Information("Cancellation token source disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing cancellation token source");
                }
                
                try
                {
                    // Dispose of services
                    _fileProcessingService?.Dispose();
                    _logger.Information("File processing service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing file processing service");
                }
                
                try
                {
                    _staThreadPool?.Dispose(); // Dispose STA thread pool
                    _logger.Information("STA thread pool disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing STA thread pool");
                }
                
                try
                {
                    _processingSemaphore?.Dispose();
                    _logger.Information("Processing semaphore disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing processing semaphore");
                }
            }

            _disposed = true;
            _logger.Information("SaveQuotesQueueService disposed");
        }

        ~SaveQuotesQueueService()
        {
            Dispose(false);
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