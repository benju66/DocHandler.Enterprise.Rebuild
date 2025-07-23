using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using DocHandler.Models;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Orchestrates document processing workflows extracted from MainViewModel
    /// Handles the complete processing pipeline from validation through completion
    /// </summary>
    public class DocumentWorkflowService : IDocumentWorkflowService
    {
        private readonly ILogger _logger;
        private readonly IOptimizedFileProcessingService _fileProcessingService;
        private readonly IConfigurationService _configService;
        private readonly IFileValidationService _fileValidationService;
        private readonly PerformanceMonitor _performanceMonitor;
        private readonly TelemetryService _telemetryService;
        
        // Active processing tracking
        private readonly Dictionary<string, DocumentProcessingContext> _activeProcessing = new();
        private readonly SemaphoreSlim _processingLock = new(1, 1);

        public DocumentWorkflowService(
            IOptimizedFileProcessingService fileProcessingService,
            IConfigurationService configService,
            IFileValidationService fileValidationService,
            PerformanceMonitor performanceMonitor,
            TelemetryService telemetryService)
        {
            _logger = Log.ForContext<DocumentWorkflowService>();
            _fileProcessingService = fileProcessingService ?? throw new ArgumentNullException(nameof(fileProcessingService));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _fileValidationService = fileValidationService ?? throw new ArgumentNullException(nameof(fileValidationService));
            _performanceMonitor = performanceMonitor ?? throw new ArgumentNullException(nameof(performanceMonitor));
            _telemetryService = telemetryService ?? throw new ArgumentNullException(nameof(telemetryService));
            
            _logger.Information("DocumentWorkflowService initialized");
        }

        public async Task<DocumentProcessingResult> ProcessDocumentAsync(
            DocumentProcessingRequest request, 
            IProgress<WorkflowProgress>? progress = null, 
            CancellationToken cancellationToken = default)
        {
            var startTime = DateTime.UtcNow;
            var context = new DocumentProcessingContext(request, startTime);
            
            try
            {
                await _processingLock.WaitAsync(cancellationToken);
                _activeProcessing[request.CorrelationId] = context;
                
                _logger.Information("Starting document processing: {CorrelationId} - {FilePath}", 
                    request.CorrelationId, request.FilePath);

                // Phase 1: Validation
                context.UpdateStatus(WorkflowStatus.Validating, "Validating document");
                progress?.Report(new WorkflowProgress 
                { 
                    CurrentOperation = "Validating document", 
                    PercentComplete = 10,
                    Status = "Validating"
                });

                var validationResult = await _fileValidationService.ValidateFileAsync(request.FilePath, cancellationToken);
                if (!validationResult.IsValid)
                {
                    return CreateFailureResult(request, $"Validation failed: {string.Join("; ", validationResult.Errors)}", startTime);
                }

                // Phase 2: Preparation
                context.UpdateStatus(WorkflowStatus.Processing, "Preparing for processing");
                progress?.Report(new WorkflowProgress 
                { 
                    CurrentOperation = "Preparing files", 
                    PercentComplete = 25,
                    Status = "Preparing"
                });

                var preparedFiles = await PrepareFilesForProcessing(new[] { request.FilePath }, cancellationToken);
                if (!preparedFiles.Any())
                {
                    return CreateFailureResult(request, "No valid files to process after preparation", startTime);
                }

                // Phase 3: Output directory preparation
                progress?.Report(new WorkflowProgress 
                { 
                    CurrentOperation = "Creating output directory", 
                    PercentComplete = 35,
                    Status = "Preparing output"
                });

                var outputDir = await PrepareOutputDirectory(request.OutputPath, cancellationToken);

                // Phase 4: Processing
                context.UpdateStatus(WorkflowStatus.Processing, "Processing document");
                progress?.Report(new WorkflowProgress 
                { 
                    CurrentOperation = "Processing document", 
                    PercentComplete = 50,
                    Status = "Processing"
                });

                var processingOptions = ExtractProcessingOptions(request.Options);
                var result = await _fileProcessingService.ProcessFiles(
                    preparedFiles, 
                    outputDir, 
                    processingOptions.ConvertOfficeToPdf).ConfigureAwait(false);

                // Phase 5: Post-processing
                progress?.Report(new WorkflowProgress 
                { 
                    CurrentOperation = "Finalizing", 
                    PercentComplete = 90,
                    Status = "Finalizing"
                });

                var finalResult = await FinalizeProcessing(request, result, outputDir, processingOptions, startTime);

                context.UpdateStatus(WorkflowStatus.Completed, "Processing completed");
                progress?.Report(new WorkflowProgress 
                { 
                    CurrentOperation = "Completed", 
                    PercentComplete = 100,
                    Status = "Completed"
                });

                _logger.Information("Document processing completed successfully: {CorrelationId}", request.CorrelationId);
                return finalResult;
            }
            catch (OperationCanceledException)
            {
                _logger.Information("Document processing cancelled: {CorrelationId}", request.CorrelationId);
                context.UpdateStatus(WorkflowStatus.Cancelled, "Processing cancelled");
                return CreateFailureResult(request, "Processing was cancelled", startTime);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error processing document: {CorrelationId} - {FilePath}", 
                    request.CorrelationId, request.FilePath);
                context.UpdateStatus(WorkflowStatus.Failed, $"Error: {ex.Message}");
                return CreateFailureResult(request, $"Processing failed: {ex.Message}", startTime);
            }
            finally
            {
                _activeProcessing.Remove(request.CorrelationId);
                _processingLock.Release();
            }
        }

        public async Task<BatchProcessingResult> ProcessDocumentsAsync(
            List<DocumentProcessingRequest> requests, 
            BatchProcessingOptions? options = null, 
            IProgress<BatchProgress>? progress = null, 
            CancellationToken cancellationToken = default)
        {
            var batchOptions = options ?? new BatchProcessingOptions();
            var results = new List<DocumentProcessingResult>();
            var startTime = DateTime.UtcNow;
            var completed = 0;
            var failed = 0;

            _logger.Information("Starting batch processing of {RequestCount} documents", requests.Count);

            var batchProgress = new BatchProgress
            {
                TotalItems = requests.Count
            };

            try
            {
                var semaphore = new SemaphoreSlim(batchOptions.MaxConcurrency, batchOptions.MaxConcurrency);
                var tasks = requests.Select(async request =>
                {
                    await semaphore.WaitAsync(cancellationToken);
                    try
                    {
                        using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                        cts.CancelAfter(batchOptions.Timeout);

                        var individualProgress = new Progress<WorkflowProgress>(p =>
                        {
                            lock (results)
                            {
                                batchProgress.CurrentItem = Path.GetFileName(request.FilePath);
                                progress?.Report(batchProgress);
                            }
                        });

                        var result = await ProcessDocumentAsync(request, individualProgress, cts.Token);
                        
                        lock (results)
                        {
                            results.Add(result);
                            if (result.Success)
                                completed++;
                            else
                                failed++;

                            batchProgress.CompletedItems = completed + failed;
                            batchProgress.FailedItems = failed;
                            progress?.Report(batchProgress);
                        }

                        if (!result.Success && batchOptions.StopOnError)
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }

                        return result;
                    }
                    finally
                    {
                        semaphore.Release();
                    }
                });

                await Task.WhenAll(tasks);
            }
            catch (OperationCanceledException)
            {
                _logger.Information("Batch processing cancelled after {CompletedCount} files", completed + failed);
            }

            var totalTime = DateTime.UtcNow - startTime;
            var batchResult = new BatchProcessingResult
            {
                TotalFiles = requests.Count,
                SuccessfulFiles = completed,
                FailedFiles = failed,
                Results = results.OrderBy(r => r.FilePath).ToList(),
                TotalProcessingTime = totalTime,
                AggregateMetrics = AggregateMetrics(results)
            };

            _logger.Information("Batch processing completed: {SuccessfulFiles}/{TotalFiles} successful in {TotalTime}",
                completed, requests.Count, totalTime);

            return batchResult;
        }

        public async Task<WorkflowStatus> GetProcessingStatusAsync(string correlationId)
        {
            await Task.Yield(); // Make it async for consistency
            
            if (_activeProcessing.TryGetValue(correlationId, out var context))
            {
                return context.Status;
            }
            
            return WorkflowStatus.Pending; // Default for unknown IDs
        }

        public async Task<bool> CancelProcessingAsync(string correlationId)
        {
            await Task.Yield(); // Make it async for consistency
            
            if (_activeProcessing.TryGetValue(correlationId, out var context))
            {
                context.CancellationTokenSource.Cancel();
                context.UpdateStatus(WorkflowStatus.Cancelled, "Cancelled by user request");
                _logger.Information("Processing cancelled for correlation ID: {CorrelationId}", correlationId);
                return true;
            }
            
            return false;
        }

        private async Task<List<string>> PrepareFilesForProcessing(string[] filePaths, CancellationToken cancellationToken)
        {
            var validFiles = new List<string>();

            foreach (var filePath in filePaths)
            {
                try
                {
                    if (!File.Exists(filePath))
                    {
                        _logger.Warning("File does not exist: {FilePath}", filePath);
                        continue;
                    }

                    // Additional validation if needed
                    if (_fileProcessingService.IsFileSupported(filePath))
                    {
                        validFiles.Add(filePath);
                    }
                    else
                    {
                        _logger.Warning("File type not supported: {FilePath}", filePath);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error preparing file: {FilePath}", filePath);
                }
            }

            return validFiles;
        }

        private async Task<string> PrepareOutputDirectory(string requestedPath, CancellationToken cancellationToken)
        {
            try
            {
                var outputDir = !string.IsNullOrEmpty(requestedPath) 
                    ? requestedPath 
                    : _configService.Config.DefaultSaveLocation;

                // Create output folder with timestamp if needed
                outputDir = _fileProcessingService.CreateOutputFolder(outputDir);

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                    _logger.Information("Created output directory: {OutputDir}", outputDir);
                }

                return outputDir;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error preparing output directory: {RequestedPath}", requestedPath);
                throw;
            }
        }

        private ProcessingOptions ExtractProcessingOptions(Dictionary<string, object> options)
        {
            var processingOptions = new ProcessingOptions();

            if (options.TryGetValue("ConvertOfficeToPdf", out var convertValue) && convertValue is bool convert)
            {
                processingOptions.ConvertOfficeToPdf = convert;
            }

            if (options.TryGetValue("OpenFolderAfterProcessing", out var openValue) && openValue is bool open)
            {
                processingOptions.OpenFolderAfterProcessing = open;
            }

            return processingOptions;
        }

        private async Task<DocumentProcessingResult> FinalizeProcessing(
            DocumentProcessingRequest request,
            dynamic processingResult,
            string outputDir,
            ProcessingOptions options,
            DateTime startTime)
        {
            try
            {
                var result = new DocumentProcessingResult
                {
                    CorrelationId = request.CorrelationId,
                    FilePath = request.FilePath,
                    OutputPath = outputDir,
                    Success = processingResult.Success,
                    ProcessingTime = DateTime.UtcNow - startTime,
                    CompletedAt = DateTime.UtcNow
                };

                if (!processingResult.Success)
                {
                    result.ErrorMessage = processingResult.ErrorMessage;
                    return result;
                }

                // Record telemetry
                if (_telemetryService != null)
                {
                    try
                    {
                        _telemetryService.TrackEvent("DocumentProcessed", new Dictionary<string, object>
                        {
                            ["FilePath"] = Path.GetFileName(request.FilePath),
                            ["ProcessingTime"] = result.ProcessingTime.TotalMilliseconds,
                            ["Success"] = result.Success
                        });
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error recording telemetry for processing completion");
                    }
                }

                // Open folder if requested
                if (options.OpenFolderAfterProcessing && Directory.Exists(outputDir))
                {
                    try
                    {
                        System.Diagnostics.Process.Start("explorer.exe", outputDir);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error opening output folder: {OutputDir}", outputDir);
                        result.Warnings.Add($"Could not open output folder: {ex.Message}");
                    }
                }

                // Collect metrics
                result.Metrics = new ProcessingMetrics
                {
                    FilesProcessed = 1,
                    BytesProcessed = new FileInfo(request.FilePath).Length,
                    MemoryUsedBytes = GC.GetTotalMemory(false),
                    CpuTime = result.ProcessingTime
                };

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error finalizing processing for: {CorrelationId}", request.CorrelationId);
                return CreateFailureResult(request, $"Finalization failed: {ex.Message}", startTime);
            }
        }

        private DocumentProcessingResult CreateFailureResult(DocumentProcessingRequest request, string errorMessage, DateTime startTime)
        {
            return new DocumentProcessingResult
            {
                CorrelationId = request.CorrelationId,
                FilePath = request.FilePath,
                Success = false,
                ErrorMessage = errorMessage,
                ProcessingTime = DateTime.UtcNow - startTime,
                CompletedAt = DateTime.UtcNow,
                Metrics = new ProcessingMetrics()
            };
        }

        private ProcessingMetrics AggregateMetrics(List<DocumentProcessingResult> results)
        {
            return new ProcessingMetrics
            {
                FilesProcessed = results.Count,
                BytesProcessed = results.Sum(r => r.Metrics.BytesProcessed),
                MemoryUsedBytes = results.Max(r => r.Metrics.MemoryUsedBytes),
                CpuTime = TimeSpan.FromTicks(results.Sum(r => r.Metrics.CpuTime.Ticks)),
                ErrorCount = results.Count(r => !r.Success),
                WarningCount = results.Sum(r => r.Warnings.Count)
            };
        }

        private class DocumentProcessingContext
        {
            public DocumentProcessingRequest Request { get; }
            public DateTime StartTime { get; }
            public WorkflowStatus Status { get; private set; }
            public string StatusMessage { get; private set; }
            public CancellationTokenSource CancellationTokenSource { get; }

            public DocumentProcessingContext(DocumentProcessingRequest request, DateTime startTime)
            {
                Request = request;
                StartTime = startTime;
                Status = WorkflowStatus.Pending;
                StatusMessage = "Queued for processing";
                CancellationTokenSource = new CancellationTokenSource();
            }

            public void UpdateStatus(WorkflowStatus status, string message)
            {
                Status = status;
                StatusMessage = message;
            }
        }

        private class ProcessingOptions
        {
            public bool ConvertOfficeToPdf { get; set; } = true;
            public bool OpenFolderAfterProcessing { get; set; } = true;
        }
    }
} 