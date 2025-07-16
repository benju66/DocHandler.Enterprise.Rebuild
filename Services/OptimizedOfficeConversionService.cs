// Folder: Services/
// File: OptimizedOfficeConversionService.cs
// Optimized Office conversion service with application pooling for improved performance
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class OptimizedOfficeConversionService : IDisposable
    {
        private readonly ILogger _logger;
        private readonly ConcurrentQueue<WordAppInstance> _wordAppPool = new();
        private readonly SemaphoreSlim _poolSemaphore;
        private readonly int _maxInstances;
        private readonly Timer _healthCheckTimer;
        private readonly Timer _memoryCleanupTimer;
        private bool _disposed;
        private bool? _officeAvailable;
        private int _conversionsCount;
        private readonly object _cleanupLock = new object();

        private readonly ConfigurationService? _configService;
        private readonly ProcessManager? _processManager;
        private readonly StaThreadPool _staThreadPool;
        private readonly OfficeInstanceTracker? _officeTracker;
        
        public OptimizedOfficeConversionService(int maxInstances = 0, ConfigurationService? configService = null, ProcessManager? processManager = null, OfficeInstanceTracker? officeTracker = null)
        {
            _logger = Log.ForContext<OptimizedOfficeConversionService>();
            _configService = configService;
            _processManager = processManager;
            _officeTracker = officeTracker;
            _staThreadPool = new StaThreadPool(1); // Single STA thread for COM operations
            
            // Conservative approach: 2-3 instances maximum based on system resources
            _maxInstances = maxInstances > 0 ? maxInstances : DetermineOptimalInstanceCount();
            _poolSemaphore = new SemaphoreSlim(_maxInstances);
            
            _logger.Information("Initializing optimized Office conversion service with {MaxInstances} Word instances", _maxInstances);
            
            // Health check timer - check every 5 minutes
            _healthCheckTimer = new Timer(PerformHealthCheck, null, 
                TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
            
            // Memory cleanup timer - cleanup every 3 minutes
            _memoryCleanupTimer = new Timer(PerformMemoryCleanup, null,
                TimeSpan.FromMinutes(3), TimeSpan.FromMinutes(3));
            
            // Initialize pool asynchronously
            Task.Run(InitializeWordAppPoolAsync);
        }

        /// <summary>
        /// Validates that the current thread is in STA apartment state for COM operations
        /// </summary>
        /// <returns>True if STA, false otherwise</returns>
        private bool ValidateSTAThread(string operation)
        {
            var apartmentState = Thread.CurrentThread.GetApartmentState();
            _logger.Debug("{Operation}: Thread {ThreadId} apartment state is {ApartmentState}", 
                operation, Thread.CurrentThread.ManagedThreadId, apartmentState);
                
            if (apartmentState != ApartmentState.STA)
            {
                _logger.Error("{Operation}: Thread is not STA - COM operations will fail", operation);
                return false;
            }
            
            return true;
        }

        private int DetermineOptimalInstanceCount()
        {
            try
            {
                // Conservative approach based on available memory and CPU cores
                using var process = Process.GetCurrentProcess();
                var availableMemoryGB = GC.GetTotalMemory(false) / (1024 * 1024 * 1024);
                var cpuCores = Environment.ProcessorCount;
                
                // Each Word instance uses ~100-200MB, be conservative
                var maxByMemory = availableMemoryGB < 4 ? 1 : (availableMemoryGB < 8 ? 2 : 3);
                var maxByCpu = Math.Max(1, cpuCores - 2); // Leave 2 cores for other processes
                
                var optimal = Math.Min(maxByMemory, maxByCpu);
                _logger.Information("Determined optimal instance count: {Count} (Memory: {MemoryGB}GB, CPU: {CpuCores})", 
                    optimal, availableMemoryGB, cpuCores);
                
                return optimal;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to determine optimal instance count, using default of 2");
                return 2;
            }
        }

        private async Task InitializeWordAppPoolAsync()
        {
            try
            {
                // Create initial Word application
                var wordApp = await CreateOptimizedWordApplicationAsync();
                if (wordApp != null)
                {
                    _wordAppPool.Enqueue(wordApp);
                    _logger.Information("Initial Word application created successfully");
                }
                else
                {
                    _logger.Warning("Failed to create initial Word application");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to initialize Word application pool");
            }
        }

        private async Task<WordAppInstance?> CreateOptimizedWordApplicationAsync()
        {
            _logger.Information("Creating optimized Word application using STA thread pool");
            
            return await _staThreadPool.ExecuteAsync(() =>
            {
                try
                {
                    // Validate we're on STA thread (StaThreadPool ensures this)
                    if (!ValidateSTAThread("CreateOptimizedWordApplication"))
                    {
                        return null;
                    }
                    
                    Type? wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null)
                    {
                        _logger.Error("OFFICE CHECK FAILED: Word.Application ProgID not found - Microsoft Word is not registered");
                        return null;
                    }
                    
                    _logger.Information("Creating Word Application COM object");
                    dynamic wordApp = Activator.CreateInstance(wordType);
                    
                    if (wordApp == null)
                    {
                        _logger.Error("OFFICE CHECK FAILED: Word Application COM object creation returned null");
                        return null;
                    }
                    
                    ComHelper.TrackComObjectCreation("WordApp", "CreateOptimizedWordApplication");
                    _logger.Information("Word Application COM object created successfully");
                    
                    // Get process ID for tracking
                    var processId = (int)wordApp.GetType().InvokeMember("ProcessID", 
                        System.Reflection.BindingFlags.GetProperty, null, wordApp, null);
                    
                    // Register with office tracker to prevent closing user instances
                    _officeTracker?.RegisterAppCreatedWordProcess(processId);
                    
                    // Apply optimizations
                    ApplyWordOptimizations(wordApp);
                    _logger.Information("Word optimizations applied successfully");
                    
                    var instance = new WordAppInstance
                    {
                        Application = wordApp,
                        ProcessId = processId,
                        CreatedAt = DateTime.UtcNow,
                        LastUsed = DateTime.UtcNow,
                        IsHealthy = true
                    };

                    _logger.Information("Created optimized Word application with PID {ProcessId}", processId);
                    return instance;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "OFFICE CHECK FAILED: Exception creating Word application");
                    return null;
                }
            });
        }

        private void ApplyWordOptimizations(dynamic wordApp)
        {
            // Critical performance optimizations
            wordApp.Visible = false;
            wordApp.DisplayAlerts = 0; // wdAlertsNone
            wordApp.ScreenUpdating = false; // Major performance boost
            wordApp.EnableEvents = false; // Disable events for performance
            wordApp.DisplayRecentFiles = false;
            wordApp.DisplayScrollBars = false;
            wordApp.DisplayStatusBar = false;
            
            // CONSERVATIVE: Only essential performance optimizations
            try
            {
                wordApp.Options.CheckGrammarAsYouType = false;
                wordApp.Options.CheckSpellingAsYouType = false;
                wordApp.Options.BackgroundSave = false;
                wordApp.Options.SaveInterval = 0; // Disable auto-save
                
                // Removed potentially unstable options:
                // - PaginationView, WPHelp, AnimateScreenMovements (UI-related)
                // - SuggestSpellingCorrections (can cause dialog issues)
            }
            catch (Exception optEx)
            {
                _logger.Warning(optEx, "Failed to set some Word optimization options");
            }
        }

        private bool IsOfficeAvailable()
        {
            if (_officeAvailable.HasValue)
            {
                _logger.Information("Office availability check (cached): {Available}", _officeAvailable.Value);
                return _officeAvailable.Value;
            }
                
            _logger.Information("=== CHECKING MICROSOFT OFFICE AVAILABILITY ===");
            
            try
            {
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    _logger.Error("OFFICE CHECK FAILED: Word.Application ProgID not found - Microsoft Word is not registered");
                    _officeAvailable = false;
                    return false;
                }
                
                _logger.Information("OFFICE CHECK: Word.Application ProgID found successfully");
                
                dynamic testApp = null;
                try
                {
                    _logger.Information("OFFICE CHECK: Attempting to create Word application instance...");
                    testApp = Activator.CreateInstance(wordType);
                    if (testApp == null)
                    {
                        _logger.Error("OFFICE CHECK FAILED: Activator.CreateInstance returned null");
                        _officeAvailable = false;
                        return false;
                    }
                    
                    ComHelper.TrackComObjectCreation("WordApp", "OptimizedOfficeAvailabilityCheck");
                    _logger.Information("OFFICE CHECK: Word application instance created successfully");
                    
                    testApp.Visible = false;
                    testApp.Quit();
                    _logger.Information("OFFICE CHECK: Word application test completed successfully");
                    
                    _officeAvailable = true;
                    _logger.Information("=== MICROSOFT OFFICE IS AVAILABLE ===");
                    return true;
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    _logger.Error(comEx, "OFFICE CHECK FAILED: COM exception - HResult={HResult}, Message={Message}", 
                        comEx.HResult, comEx.Message);
                    _officeAvailable = false;
                    return false;
                }
                finally
                {
                    if (testApp != null)
                    {
                        try
                        {
                            ComHelper.SafeReleaseComObject(testApp, "WordApp", "OptimizedOfficeAvailabilityCheck");
                        }
                        catch (Exception disposeEx)
                        {
                            _logger.Warning(disposeEx, "Error disposing Word application test instance");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "OFFICE CHECK FAILED: Unexpected error - {ExceptionType}: {Message}", 
                    ex.GetType().Name, ex.Message);
            }
            
            _logger.Error("=== MICROSOFT OFFICE IS NOT AVAILABLE ===");
            _officeAvailable = false;
            return false;
        }

        public async Task<ConversionResult> ConvertWordToPdf(string inputPath, string outputPath)
        {
            if (!IsOfficeAvailable())
            {
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = "Microsoft Office is not installed or accessible. Please install Microsoft Office to convert Word documents to PDF."
                };
            }

            await _poolSemaphore.WaitAsync();
            WordAppInstance? wordInstance = null;
            
            try
            {
                wordInstance = await GetHealthyWordAppInstance();
                if (wordInstance == null)
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Could not obtain healthy Word application instance"
                    };
                }

                var result = await ConvertWordToPdfWithInstance(wordInstance, inputPath, outputPath);
                
                // Update usage statistics
                wordInstance.LastUsed = DateTime.UtcNow;
                Interlocked.Increment(ref _conversionsCount);
                
                return result;
            }
            finally
            {
                // Always return Word instance to pool if it's still healthy
                if (wordInstance != null && wordInstance.IsHealthy)
                {
                    _wordAppPool.Enqueue(wordInstance);
                }
                
                _poolSemaphore.Release();
            }
        }

        private async Task<WordAppInstance?> GetHealthyWordAppInstance()
        {
            const int maxRetries = 3;
            
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                // Try to get instance from pool first
                if (_wordAppPool.TryDequeue(out var wordInstance))
                {
                    if (await IsWordInstanceHealthy(wordInstance))
                    {
                        return wordInstance;
                    }
                    else
                    {
                        _logger.Warning("Unhealthy Word instance detected, disposing and creating new one");
                        await DisposeWordInstance(wordInstance);
                    }
                }
                
                // Create new instance if pool is empty or instances are unhealthy
                if (_wordAppPool.Count < _maxInstances)
                {
                    var newInstance = await CreateOptimizedWordApplicationAsync();
                    if (newInstance != null)
                    {
                        return newInstance;
                    }
                }
                
                // Wait before retry
                if (attempt < maxRetries - 1)
                {
                    await Task.Delay(1000);
                }
            }
            
            _logger.Error("Could not obtain healthy Word application after {MaxRetries} attempts", maxRetries);
            return null;
        }

        private async Task<bool> IsWordInstanceHealthy(WordAppInstance instance)
        {
            try
            {
                // Check if process is still running
                if (!IsProcessRunning(instance.ProcessId))
                {
                    instance.IsHealthy = false;
                    return false;
                }

                // Simple health check - try to access Word application properties
                var isHealthy = await _staThreadPool.ExecuteAsync(async () =>
                {
                    ValidateSTAThread("WordInstanceHealthCheck");
                    
                    var _ = instance.Application.Version;
                    var documentCount = instance.Application.Documents.Count;
                    return true;
                });

                return isHealthy;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Word instance health check failed for PID {ProcessId}", instance.ProcessId);
                instance.IsHealthy = false;
                return false;
            }
        }

        private bool IsProcessRunning(int processId)
        {
            try
            {
                var process = Process.GetProcessById(processId);
                return !process.HasExited;
            }
            catch (ArgumentException)
            {
                // Process doesn't exist
                return false;
            }
        }

        private async Task<ConversionResult> ConvertWordToPdfWithInstance(WordAppInstance wordInstance, string inputPath, string outputPath)
        {
            var timeoutSeconds = _configService?.Config.ConversionTimeoutSeconds ?? 30;
            
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
            
            try
            {
                // SIMPLIFIED: We're already on STA thread from queue processing
                if (!ValidateSTAThread("ConvertWordToPdfWithInstance"))
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Thread apartment state is not STA - COM operations will fail"
                    };
                }
                
                var result = new ConversionResult();
                dynamic? doc = null;
                var stopwatch = Stopwatch.StartNew();
            
                try
                {
                    // Handle network paths by copying locally if needed
                    var workingInputPath = inputPath;
                    string? tempFilePath = null;
                    
                    if (IsNetworkPath(inputPath))
                    {
                        tempFilePath = Path.Combine(Path.GetTempPath(), $"DocHandler_{Guid.NewGuid()}.docx");
                        File.Copy(inputPath, tempFilePath, true);
                        workingInputPath = tempFilePath;
                        _logger.Debug("Network path detected, using local copy: {TempPath}", tempFilePath);
                    }

                    try
                    {
                        // Open document with performance-optimized settings
                        doc = wordInstance.Application.Documents.Open(
                            workingInputPath,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            Repair: false,
                            ShowRepairs: false,
                            OpenAndRepair: false,
                            NoEncodingDialog: true,
                            Revert: false
                        );
                        ComHelper.TrackComObjectCreation("Document", "OptimizedConversion");

                        // Convert to PDF with optimized settings for speed
                        doc.SaveAs2(
                            outputPath,
                            FileFormat: 17, // wdFormatPDF
                            EmbedTrueTypeFonts: false, // Faster
                            SaveNativePictureFormat: false, // Faster
                            SaveFormsData: false, // Faster
                            CompressLevel: 0, // Fastest compression
                            UseDocumentImageQuality: false, // Faster
                            IncludeDocProps: false, // Faster
                            KeepIRM: false, // Faster
                            CreateBookmarks: false, // Faster
                            DocStructureTags: false, // Faster
                            BitmapMissingFonts: false // Faster
                        );

                        result.Success = true;
                        result.OutputPath = outputPath;
                        
                        stopwatch.Stop();
                        _logger.Information("Converted {File} in {ElapsedMs}ms using PID {ProcessId}", 
                            Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds, wordInstance.ProcessId);
                    }
                    finally
                    {
                        // Clean up temporary file
                        if (tempFilePath != null && File.Exists(tempFilePath))
                        {
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch (Exception ex)
                            {
                                _logger.Warning(ex, "Failed to delete temporary file: {TempPath}", tempFilePath);
                            }
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    _logger.Error(comEx, "COM error during conversion: {HResult}", comEx.HResult);
                    
                    // Handle specific COM errors
                    if (comEx.HResult == -2147221164) // RPC_E_CANTCALLOUT_ININPUTSYNCCALL
                    {
                        result.ErrorMessage = "COM threading error - Word instance will be recreated";
                        wordInstance.IsHealthy = false; // Mark instance as unhealthy
                    }
                    else
                    {
                        result.ErrorMessage = $"COM error: {comEx.Message}";
                    }
                    
                    result.Success = false;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to convert Word document: {Path}", inputPath);
                    result.Success = false;
                    result.ErrorMessage = $"Conversion failed: {ex.Message}";
                }
                finally
                {
                    // Clean up document
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                            ComHelper.SafeReleaseComObject(doc, "Document", "OptimizedConversion");
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word document");
                        }
                    }
                }

                return result;
            }
            catch (OperationCanceledException)
            {
                _logger.Error("Word conversion timed out after {TimeoutSeconds} seconds for file: {InputPath}", 
                    timeoutSeconds, inputPath);
                
                // Mark instance as unhealthy so it gets recreated
                wordInstance.IsHealthy = false;
                
                // Clean up orphaned processes
                try
                {
                    _processManager?.KillOrphanedWordProcesses();
                    _logger.Information("Cleaned up orphaned Word processes after timeout");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error cleaning up orphaned processes after timeout");
                }
                
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Word conversion timed out after {timeoutSeconds} seconds"
                };
            }
        }

        private bool IsNetworkPath(string path)
        {
            try
            {
                return path.StartsWith(@"\\") || 
                       (Path.IsPathRooted(path) && new DriveInfo(path).DriveType == DriveType.Network);
            }
            catch
            {
                return false;
            }
        }

        private void PerformHealthCheck(object? state)
        {
            try
            {
                var instances = new List<WordAppInstance>();
                
                // Collect all instances from pool
                while (_wordAppPool.TryDequeue(out var instance))
                {
                    instances.Add(instance);
                }
                
                // Check health of each instance
                foreach (var instance in instances)
                {
                    if (IsWordInstanceHealthy(instance).Result)
                    {
                        _wordAppPool.Enqueue(instance);
                    }
                    else
                    {
                        _logger.Information("Disposing unhealthy Word instance PID {ProcessId}", instance.ProcessId);
                        _ = DisposeWordInstance(instance);
                    }
                }
                
                _logger.Debug("Health check completed. Pool size: {PoolSize}", _wordAppPool.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during health check");
            }
        }

        private void PerformMemoryCleanup(object? state)
        {
            lock (_cleanupLock)
            {
                if (_conversionsCount > 5) // Only cleanup after significant activity
                {
                    _logger.Debug("Performing memory cleanup after {Count} conversions", _conversionsCount);
                    
                    // Force cleanup of COM objects
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    
                    _conversionsCount = 0;
                }
            }
        }

        private async Task DisposeWordInstance(WordAppInstance instance)
        {
            try
            {
                if (instance.Application != null)
                {
                    // Unregister from office tracker first
                    _officeTracker?.UnregisterAppCreatedWordProcess(instance.ProcessId);
                    
                    await _staThreadPool.ExecuteAsync(async () =>
                    {
                        try
                        {
                            // Close all documents
                            var documents = instance.Application.Documents;
                            if (documents != null && documents.Count > 0)
                            {
                                foreach (dynamic doc in documents)
                                {
                                    try
                                    {
                                        doc.Close(SaveChanges: false);
                                        ComHelper.SafeReleaseComObject(doc, "Document", "DisposeWordInstance");
                                    }
                                    catch { }
                                }
                                ComHelper.SafeReleaseComObject(documents, "Documents", "DisposeWordInstance");
                            }
                            
                            // Quit Word application
                            instance.Application.Quit();
                            ComHelper.SafeReleaseComObject(instance.Application, "WordApp", "DisposeWordInstance");
                            
                            _logger.Information("Disposed Word application PID {ProcessId}", instance.ProcessId);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error during Word instance disposal");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to dispose Word instance PID {ProcessId}", instance.ProcessId);
            }
        }

        public bool IsOfficeInstalled()
        {
            return IsOfficeAvailable();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _healthCheckTimer?.Dispose();
                    _memoryCleanupTimer?.Dispose();
                    _poolSemaphore?.Dispose();
                }
                
                // Clean up all Word instances in pool
                while (_wordAppPool.TryDequeue(out var instance))
                {
                    _ = DisposeWordInstance(instance);
                }
                
                // Dispose STA thread pool
                _staThreadPool?.Dispose();
                
                // Force garbage collection
                ComHelper.ForceComCleanup("OptimizedOfficeConversionService");
                
                // Log final COM statistics
                ComHelper.LogComObjectStats();
                
                _disposed = true;
                _logger.Information("Optimized Office conversion service disposed");
            }
        }

        ~OptimizedOfficeConversionService()
        {
            Dispose(false);
        }
    }

    public class WordAppInstance
    {
        public dynamic? Application { get; set; }
        public int ProcessId { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime LastUsed { get; set; }
        public bool IsHealthy { get; set; }
    }
} 