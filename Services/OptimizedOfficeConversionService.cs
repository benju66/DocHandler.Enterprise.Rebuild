// Folder: Services/
// File: OptimizedOfficeConversionService.cs
// Optimized Office conversion service with application pooling for improved performance
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

        // Windows API imports for safe process ID retrieval
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);
        
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

        /// <summary>
        /// Safely gets Word process ID using window handle approach instead of ProcessID property
        /// </summary>
        private int GetWordProcessIdSafely(dynamic wordApp)
        {
            try
            {
                IntPtr hwnd = IntPtr.Zero;
                
                // Method 1: Try to get application window handle (Word 2010+)
                try
                {
                    hwnd = new IntPtr((int)wordApp.Hwnd);
                }
                catch (Exception ex)
                {
                    _logger.Debug("Could not get Hwnd property: {Message}", ex.Message);
                    
                    // Method 2: Try ActiveWindow.Hwnd (if document is open)
                    try
                    {
                        if (wordApp.ActiveWindow != null)
                        {
                            hwnd = new IntPtr((int)wordApp.ActiveWindow.Hwnd);
                        }
                    }
                    catch (Exception ex2)
                    {
                        _logger.Debug("Could not get ActiveWindow.Hwnd: {Message}", ex2.Message);
                    }
                }
                
                // If we got a valid window handle, get the process ID
                if (hwnd != IntPtr.Zero && IsWindow(hwnd))
                {
                    uint processId;
                    if (GetWindowThreadProcessId(hwnd, out processId) != 0)
                    {
                        _logger.Debug("Successfully retrieved process ID {ProcessId} via window handle", processId);
                        return (int)processId;
                    }
                }
                
                _logger.Debug("Could not retrieve process ID - window handle method failed");
                return 0;
            }
            catch (Exception ex)
            {
                _logger.Debug("Error getting Word process ID safely: {Message}", ex.Message);
                return 0;
            }
        }

        /// <summary>
        /// Safely sets Word application properties with DISP_E_UNKNOWNNAME error handling
        /// </summary>
        private bool SafeSetProperty(dynamic obj, string propertyName, object value)
        {
            try
            {
                obj.GetType().InvokeMember(propertyName, 
                    System.Reflection.BindingFlags.SetProperty, null, obj, new[] { value });
                return true;
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
            {
                _logger.Debug("Property {Property} not found (DISP_E_UNKNOWNNAME) - version compatibility issue", propertyName);
                return false;
            }
            catch (Exception ex)
            {
                _logger.Debug("Failed to set property {Property}: {Message}", propertyName, ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Safely opens a Word document with fallback parameter sets for version compatibility
        /// </summary>
        private dynamic OpenDocumentSafely(dynamic wordApp, string filePath)
        {
            dynamic documents = null;
            try
            {
                documents = wordApp.Documents;
                ComHelper.TrackComObjectCreation("Documents", "OpenDocument");
                
                // Try with full parameters first (Word 2010+)
                try
                {
                    return documents.Open(
                        filePath,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        Repair: false,
                        ShowRepairs: false,
                        OpenAndRepair: false,
                        NoEncodingDialog: true,
                        Revert: false
                    );
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                {
                    _logger.Debug("Extended Open parameters not supported, using basic Open");
                    // Fallback to basic Open (Word 2007+)
                    return documents.Open(filePath, ReadOnly: true);
                }
            }
            finally
            {
                // Always release the Documents collection, but NOT the returned document
                if (documents != null)
                {
                    ComHelper.SafeReleaseComObject(documents, "Documents", "OpenDocument");
                }
            }
        }

        /// <summary>
        /// Safely saves document as PDF with fallback methods for version compatibility
        /// </summary>
        private bool SaveAsPdfSafely(dynamic doc, string outputPath)
        {
            try
            {
                // Try SaveAs2 with full parameters (Word 2010+)
                doc.SaveAs2(
                    outputPath,
                    FileFormat: 17, // wdFormatPDF
                    EmbedTrueTypeFonts: false,
                    SaveNativePictureFormat: false,
                    SaveFormsData: false,
                    CompressLevel: 0,
                    UseDocumentImageQuality: false,
                    IncludeDocProps: false,
                    KeepIRM: false,
                    CreateBookmarks: false,
                    DocStructureTags: false,
                    BitmapMissingFonts: false
                );
                _logger.Debug("Document saved with extended SaveAs2 parameters");
                return true;
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
            {
                _logger.Debug("SaveAs2 with extended parameters not available, trying basic SaveAs2");
                try
                {
                    // Fallback to basic SaveAs2
                    doc.SaveAs2(outputPath, FileFormat: 17);
                    _logger.Debug("Document saved with basic SaveAs2");
                    return true;
                }
                catch (COMException ex2) when (ex2.HResult == unchecked((int)0x80020006))
                {
                    _logger.Debug("SaveAs2 not available, trying SaveAs");
                    // Final fallback to SaveAs (Word 2007)
                    doc.SaveAs(outputPath, FileFormat: 17);
                    _logger.Debug("Document saved with legacy SaveAs");
                    return true;
                }
            }
        }

        /// <summary>
        /// Gets a user-friendly error message for COM exceptions
        /// </summary>
        private string GetCOMErrorMessage(COMException ex)
        {
            return ex.HResult switch
            {
                unchecked((int)0x80020006) => "Property or method not found (DISP_E_UNKNOWNNAME) - Word version incompatibility",
                unchecked((int)0x800A11FD) => "Document window is not active",
                -2147221164 => "COM threading error (RPC_E_CANTCALLOUT_ININPUTSYNCCALL)",
                unchecked((int)0x80010108) => "COM object has been disconnected from its underlying RPC server",
                unchecked((int)0x800706BE) => "Remote procedure call failed",
                _ => $"COM error: {ex.Message} (HResult: {ex.HResult:X8})"
            };
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
                    
                    // ProcessID tracking is optional - if we can't get it, we continue without it
                    int processId = 0;
                    try
                    {
                        // Some Word versions don't expose ProcessID through COM
                        processId = GetWordProcessIdSafely(wordApp);
                        _logger.Information("Word ProcessID accessed: {ProcessId}", processId);
                    }
                    catch (Exception ex)
                    {
                        _logger.Debug("ProcessID not available (expected on some Word versions): {Message}", ex.Message);
                        // Continue without process tracking - not critical
                    }
                    
                    // Only register if we got a valid ProcessID
                    if (processId > 0 && _officeTracker != null)
                    {
                        _officeTracker.RegisterAppCreatedWordProcess(processId);
                    }
                    
                    // Apply optimizations
                    ApplyWordOptimizations(wordApp);
                    _logger.Information("Word optimizations applied successfully");
                    
                    var instance = new WordAppInstance
                    {
                        Application = wordApp,
                        ProcessId = processId, // Will be 0 if not available
                        CreatedAt = DateTime.UtcNow,
                        LastUsed = DateTime.UtcNow,
                        IsHealthy = true
                    };

                    _logger.Information("Created optimized Word application");
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
            // Critical performance optimizations - these should work on all versions
            SafeSetProperty(wordApp, "Visible", false);
            SafeSetProperty(wordApp, "DisplayAlerts", 0); // wdAlertsNone
            SafeSetProperty(wordApp, "ScreenUpdating", false); // Major performance boost
            SafeSetProperty(wordApp, "EnableEvents", false); // Disable events for performance
            SafeSetProperty(wordApp, "DisplayRecentFiles", false);
            SafeSetProperty(wordApp, "DisplayScrollBars", false);
            SafeSetProperty(wordApp, "DisplayStatusBar", false);
            
            // CONSERVATIVE: Only essential performance optimizations with error handling
            try
            {
                if (wordApp.Options != null)
                {
                    SafeSetProperty(wordApp.Options, "CheckGrammarAsYouType", false);
                    SafeSetProperty(wordApp.Options, "CheckSpellingAsYouType", false);
                    SafeSetProperty(wordApp.Options, "BackgroundSave", false);
                    SafeSetProperty(wordApp.Options, "SaveInterval", 0); // Disable auto-save
                }
            }
            catch (Exception optEx)
            {
                _logger.Debug("Some Word Options properties not available: {Message}", optEx.Message);
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

            // Always use synchronous conversion when on STA thread (queue processing)
            // This avoids the nested STA thread pool issue
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                _logger.Debug("On STA thread, using direct synchronous conversion (avoiding nested thread pools)");
                return ConvertWordToPdfSync(inputPath, outputPath);
            }

            // Only use the pooled approach for non-STA thread calls (rare case)
            _logger.Debug("On non-STA thread, using pooled conversion approach");
            return await ConvertWordToPdfWithPooling(inputPath, outputPath);
        }

        /// <summary>
        /// Synchronous Word to PDF conversion for use when already on STA thread
        /// </summary>
        private ConversionResult ConvertWordToPdfSync(string inputPath, string outputPath)
        {
            var result = new ConversionResult();
            dynamic? wordApp = null;
            dynamic? doc = null;
            var stopwatch = Stopwatch.StartNew();

            try
            {
                _logger.Information("Converting Word to PDF (sync): {InputPath} -> {OutputPath}", inputPath, outputPath);
                
                // Verify STA thread state
                var apartmentState = Thread.CurrentThread.GetApartmentState();
                _logger.Debug("ConvertWordToPdfSync: Running on {ApartmentState} thread {ThreadId}", 
                    apartmentState, Thread.CurrentThread.ManagedThreadId);
                
                if (apartmentState != ApartmentState.STA)
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = $"Thread must be STA for COM operations. Current state: {apartmentState}"
                    };
                }
                
                // Create Word application directly on current STA thread
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    _logger.Error("ConvertWordToPdfSync: Word.Application ProgID not found");
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Microsoft Word is not installed or accessible - ProgID not found."
                    };
                }

                _logger.Debug("ConvertWordToPdfSync: Creating Word application instance");
                
                try
                {
                    wordApp = Activator.CreateInstance(wordType);
                }
                catch (COMException comEx)
                {
                    _logger.Error(comEx, "ConvertWordToPdfSync: COM exception creating Word application - HResult={HResult}", comEx.HResult);
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = $"COM error creating Word application: {comEx.Message} (HResult: {comEx.HResult:X8})"
                    };
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "ConvertWordToPdfSync: Exception creating Word application");
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = $"Error creating Word application: {ex.Message}"
                    };
                }
                
                if (wordApp == null)
                {
                    _logger.Error("ConvertWordToPdfSync: Word application creation returned null");
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to create Word application instance - Activator returned null."
                    };
                }
                
                ComHelper.TrackComObjectCreation("WordApp", "OptimizedSyncConversion");
                _logger.Debug("ConvertWordToPdfSync: Word application created successfully");
                
                // Configure Word for optimal conversion performance
                try
                {
                    wordApp.Visible = false;
                    wordApp.DisplayAlerts = 0; // wdAlertsNone
                    wordApp.ScreenUpdating = false;
                    _logger.Debug("ConvertWordToPdfSync: Word application configured for conversion");
                }
                catch (Exception configEx)
                {
                    _logger.Warning(configEx, "ConvertWordToPdfSync: Failed to configure Word application - continuing anyway");
                }

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
                    doc = OpenDocumentSafely(wordApp, workingInputPath);
                    ComHelper.TrackComObjectCreation("Document", "OptimizedSyncConversion");

                    // Convert to PDF with optimized settings for speed
                    if (SaveAsPdfSafely(doc, outputPath))
                    {
                        result.Success = true;
                        result.OutputPath = outputPath;
                        
                        stopwatch.Stop();
                        _logger.Information("Successfully converted Word to PDF in {ElapsedMs}ms (sync)", stopwatch.ElapsedMilliseconds);
                    }
                    else
                    {
                        _logger.Error("Failed to save document as PDF (sync)");
                        result.Success = false;
                        result.ErrorMessage = "Failed to save document as PDF.";
                    }
                }
                finally
                {
                    // Clean up temporary file
                    if (tempFilePath != null && File.Exists(tempFilePath))
                    {
                        try { File.Delete(tempFilePath); } catch { }
                    }
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                _logger.Error(comEx, "COM error during synchronous conversion: {HResult}", comEx.HResult);
                result.Success = false;
                result.ErrorMessage = GetCOMErrorMessage(comEx);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to convert Word document (sync): {Path}", inputPath);
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
                        ComHelper.SafeReleaseComObject(doc, "Document", "OptimizedSyncConversion");
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error closing Word document (sync)");
                    }
                }

                // Clean up Word application
                if (wordApp != null)
                {
                    try
                    {
                        wordApp.Quit(SaveChanges: false);
                        ComHelper.SafeReleaseComObject(wordApp, "WordApp", "OptimizedSyncConversion");
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error closing Word application (sync)");
                    }
                }

                // Force COM cleanup
                ComHelper.ForceComCleanup("OptimizedSyncConversion");
            }

            return result;
        }

        /// <summary>
        /// Pooled conversion approach for non-STA thread calls
        /// </summary>
        private async Task<ConversionResult> ConvertWordToPdfWithPooling(string inputPath, string outputPath)
        {
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
                // Check process if we have a valid ProcessId
                if (instance.ProcessId > 0 && !IsProcessRunning(instance.ProcessId))
                {
                    instance.IsHealthy = false;
                    return false;
                }

                // Comprehensive health check - try multiple properties to verify Word is responsive
                var isHealthy = await _staThreadPool.ExecuteAsync(() =>
                {
                    ValidateSTAThread("WordInstanceHealthCheck");
                    
                    try
                    {
                        // Check 1: Can access application object
                        if (instance.Application == null)
                            return false;
                        
                        // Check 2: Try to access Documents collection
                        using (var comScope = new ComResourceScope())
                        {
                            var documents = comScope.GetDocuments(instance.Application, "HealthCheck");
                            var docCount = documents.Count;
                        } // documents collection automatically released here
                        
                        // Check 3: Try to access Version (basic property that should always work)
                        try
                        {
                            var version = instance.Application.Version;
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                        {
                            // Version property not available - but that's okay, continue
                            _logger.Debug("Version property not available during health check");
                        }
                        
                        // Check 4: Verify application is responding (not in a hung state)
                        try
                        {
                            var visible = instance.Application.Visible;
                        }
                        catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                        {
                            // Visible property not available - but that's okay, continue
                            _logger.Debug("Visible property not available during health check");
                        }
                        
                        return true;
                    }
                    catch (COMException comEx)
                    {
                        _logger.Debug("Word instance COM health check failed: {Message}", GetCOMErrorMessage(comEx));
                        return false;
                    }
                    catch (Exception ex)
                    {
                        _logger.Debug("Word instance health check failed: {Message}", ex.Message);
                        return false;
                    }
                });

                if (!isHealthy)
                {
                    instance.IsHealthy = false;
                }

                return isHealthy;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Word instance health check failed");
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
                // Log current thread state for debugging, but don't fail if already on STA thread
                var apartmentState = Thread.CurrentThread.GetApartmentState();
                _logger.Debug("ConvertWordToPdfWithInstance: Thread {ThreadId} apartment state is {ApartmentState}", 
                    Thread.CurrentThread.ManagedThreadId, apartmentState);
                
                if (apartmentState != ApartmentState.STA)
                {
                    _logger.Error("ConvertWordToPdfWithInstance: Thread is not STA - COM operations will fail");
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
                        doc = OpenDocumentSafely(wordInstance.Application, workingInputPath);
                        ComHelper.TrackComObjectCreation("Document", "OptimizedConversion");

                        // Convert to PDF with optimized settings for speed
                        if (SaveAsPdfSafely(doc, outputPath))
                        {
                            result.Success = true;
                            result.OutputPath = outputPath;
                            
                            stopwatch.Stop();
                            _logger.Information("Converted {File} in {ElapsedMs}ms", 
                                Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds);
                        }
                        else
                        {
                            _logger.Error("Failed to save document as PDF (pooled)");
                            result.Success = false;
                            result.ErrorMessage = "Failed to save document as PDF.";
                        }
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
                        result.ErrorMessage = GetCOMErrorMessage(comEx);
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
                        _logger.Information("Disposing unhealthy Word instance");
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
                    // Only unregister if we have a valid ProcessId
                    if (instance.ProcessId > 0 && _officeTracker != null)
                    {
                        _officeTracker.UnregisterAppCreatedWordProcess(instance.ProcessId);
                    }
                    
                    await _staThreadPool.ExecuteAsync(async () =>
                    {
                        try
                        {
                            // Close all documents
                            using (var comScope = new ComResourceScope())
                            {
                                var documents = comScope.GetDocuments(instance.Application, "DisposeWordInstance");
                                if (documents != null && documents.Count > 0)
                                {
                                    foreach (dynamic doc in documents)
                                    {
                                        try
                                        {
                                            doc.Close(SaveChanges: false);
                                            comScope.Track(doc, "Document", "DisposeWordInstance");
                                        }
                                        catch { }
                                    }
                                }
                            } // documents collection automatically released here
                            
                            // Quit Word application
                            instance.Application.Quit();
                            ComHelper.SafeReleaseComObject(instance.Application, "WordApp", "DisposeWordInstance");
                            
                            _logger.Information("Disposed Word application");
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
                _logger.Warning(ex, "Failed to dispose Word instance");
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