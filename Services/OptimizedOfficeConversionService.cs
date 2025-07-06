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

        public OptimizedOfficeConversionService(int maxInstances = 0)
        {
            _logger = Log.ForContext<OptimizedOfficeConversionService>();
            
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
            return await Task.Run(() =>
            {
                try
                {
                    // Set COM apartment state for this thread
                    Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                    
                    Type? wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null)
                    {
                        _logger.Error("Word.Application ProgID not found");
                        return null;
                    }

                    dynamic wordApp = Activator.CreateInstance(wordType);
                    if (wordApp == null)
                    {
                        _logger.Error("Failed to create Word application instance");
                        return null;
                    }

                    // Critical performance optimizations
                    wordApp.Visible = false;
                    wordApp.DisplayAlerts = 0; // wdAlertsNone
                    wordApp.ScreenUpdating = false; // Major performance boost
                    wordApp.EnableEvents = false; // Disable events for performance
                    wordApp.DisplayRecentFiles = false;
                    wordApp.DisplayScrollBars = false;
                    wordApp.DisplayStatusBar = false;
                    
                    // Disable automatic features that slow down conversion
                    try
                    {
                        wordApp.Options.CheckGrammarAsYouType = false;
                        wordApp.Options.CheckSpellingAsYouType = false;
                        wordApp.Options.SuggestSpellingCorrections = false;
                        wordApp.Options.BackgroundSave = false;
                        wordApp.Options.SaveInterval = 0; // Disable auto-save
                        wordApp.Options.PaginationView = false;
                        wordApp.Options.WPHelp = false;
                        wordApp.Options.AnimateScreenMovements = false;
                    }
                    catch (Exception optEx)
                    {
                        _logger.Warning(optEx, "Failed to set some Word optimization options");
                    }

                    // Get process ID for tracking
                    var processId = (int)wordApp.GetType().InvokeMember("ProcessID", 
                        System.Reflection.BindingFlags.GetProperty, null, wordApp, null);

                    var instance = new WordAppInstance
                    {
                        Application = wordApp,
                        ProcessId = processId,
                        CreatedAt = DateTime.UtcNow,
                        LastUsed = DateTime.UtcNow,
                        IsHealthy = true
                    };

                    _logger.Debug("Created optimized Word application with PID {ProcessId}", processId);
                    return instance;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to create optimized Word application");
                    return null;
                }
            });
        }

        private bool IsOfficeAvailable()
        {
            if (_officeAvailable.HasValue)
                return _officeAvailable.Value;
                
            try
            {
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType != null)
                {
                    dynamic testApp = null;
                    try
                    {
                        testApp = Activator.CreateInstance(wordType);
                        testApp.Visible = false;
                        testApp.Quit();
                        _officeAvailable = true;
                        return true;
                    }
                    finally
                    {
                        if (testApp != null)
                        {
                            try
                            {
                                Marshal.ReleaseComObject(testApp);
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Microsoft Office is not available");
            }
            
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
                await Task.Run(() =>
                {
                    var _ = instance.Application.Version;
                    var documentCount = instance.Application.Documents.Count;
                });

                return true;
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
            return await Task.Run(() =>
            {
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
                            Marshal.ReleaseComObject(doc);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word document");
                        }
                    }
                }

                return result;
            });
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
                    await Task.Run(() =>
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
                                        Marshal.ReleaseComObject(doc);
                                    }
                                    catch { }
                                }
                                Marshal.ReleaseComObject(documents);
                            }
                            
                            // Quit Word application
                            instance.Application.Quit();
                            Marshal.ReleaseComObject(instance.Application);
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
                
                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
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