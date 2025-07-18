using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Maintains a single Word application instance for the session to improve performance
    /// </summary>
    public class SessionAwareOfficeService : IDisposable
    {
        private readonly ILogger _logger;
        private dynamic _wordApp;
        private DateTime _lastUsed;
        private Timer _idleTimer;
        private readonly object _wordLock = new object();
        private readonly TimeSpan _idleTimeout = TimeSpan.FromMinutes(5);
        private bool _disposed;
        private bool? _officeAvailable;
        private DateTime _lastHealthCheck = DateTime.Now;
        private readonly TimeSpan _healthCheckInterval = TimeSpan.FromMinutes(5);
        private readonly ConversionCircuitBreaker _circuitBreaker;

        // Windows API imports for safe process ID retrieval
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);
        
        public SessionAwareOfficeService()
        {
            _logger = Log.ForContext<SessionAwareOfficeService>();
            _circuitBreaker = new ConversionCircuitBreaker();
            _logger.Information("Initializing session-aware Office service with circuit breaker protection");
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
        
        private bool IsOfficeAvailable()
        {
            if (_officeAvailable.HasValue)
                return _officeAvailable.Value;
                
            try
            {
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType != null)
                {
                    _officeAvailable = true;
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Microsoft Office is not available");
            }
            
            _officeAvailable = false;
            return false;
        }
        
        private dynamic GetOrCreateWordApp()
        {
            lock (_wordLock)
            {
                _lastUsed = DateTime.Now;
                
                if (_wordApp == null)
                {
                    try
                    {
                        // Validate STA thread before COM operations
                        var apartmentState = Thread.CurrentThread.GetApartmentState();
                        if (apartmentState != ApartmentState.STA)
                        {
                            _logger.Error("GetOrCreateWordApp: Thread is not STA ({ApartmentState}) - COM operations may fail", apartmentState);
                            throw new InvalidOperationException($"Thread must be STA for COM operations. Current state: {apartmentState}");
                        }
                        
                        _logger.Information("Creating new Word application instance for session");
                        Type wordType = Type.GetTypeFromProgID("Word.Application");
                        if (wordType == null)
                        {
                            throw new InvalidOperationException("Word.Application ProgID not found");
                        }
                        
                        _wordApp = Activator.CreateInstance(wordType);
                        ComHelper.TrackComObjectCreation("WordApp", "SessionAwareOfficeService");
                        
                        // Apply optimizations safely
                        ApplyWordOptimizations(_wordApp);
                        
                        // Get process ID safely for tracking
                        try
                        {
                            var processId = GetWordProcessIdSafely(_wordApp);
                            if (processId > 0)
                            {
                                _logger.Information("Word application created with process ID: {ProcessId}", processId);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Debug("Could not get process ID: {Message}", ex.Message);
                        }
                        
                        _lastHealthCheck = DateTime.Now;
                        _logger.Information("Word application initialized for session");
                        
                        // Set up idle timer
                        _idleTimer = new Timer(CheckIdleTimeout, null, _idleTimeout, _idleTimeout);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to create Word application");
                        throw;
                    }
                }
                
                return _wordApp;
            }
        }

        private void ApplyWordOptimizations(dynamic wordApp)
        {
            try
            {
                // Only set the most essential properties
                wordApp.Visible = false;
                wordApp.DisplayAlerts = 0; // wdAlertsNone
                // Skip other properties during warm-up to avoid COM object creation
            }
            catch (Exception ex)
            {
                _logger.Debug("Error applying Word optimizations: {Message}", ex.Message);
                // Continue anyway - these are just optimizations
            }
            
            // Advanced optimizations with error handling
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
                // Try SaveAs2 with optimized parameters (Word 2010+)
                doc.SaveAs2(
                    outputPath, 
                    FileFormat: 17, // wdFormatPDF
                    EmbedTrueTypeFonts: false,
                    SaveNativePictureFormat: false,
                    SaveFormsData: false,
                    CompressLevel: 0
                );
                _logger.Debug("Document saved with optimized SaveAs2 parameters");
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

        private void CheckIdleTimeout(object state)
        {
            lock (_wordLock)
            {
                if (_wordApp != null && DateTime.Now - _lastUsed > _idleTimeout)
                {
                    _logger.Information("Word application idle for {Minutes} minutes, disposing", _idleTimeout.TotalMinutes);
                    DisposeWordApp();
                }
            }
        }

        private bool IsWordHealthy()
        {
            try
            {
                if (_wordApp == null) return false;
                
                // Don't access properties during health check - they create COM objects
                // Just check if the reference is not null
                return true;
            }
            catch
            {
                _logger.Warning("Word health check failed - application may have crashed");
                return false;
            }
        }

        private void EnsureWordHealthy()
        {
            if (DateTime.Now - _lastHealthCheck > _healthCheckInterval)
            {
                if (!IsWordHealthy())
                {
                    _logger.Warning("Word unhealthy - reinitializing");
                    DisposeWordApp();
                }
                _lastHealthCheck = DateTime.Now;
            }
        }
        
        public async Task<ConversionResult> ConvertWordToPdf(string inputPath, string outputPath)
        {
            if (!IsOfficeAvailable())
            {
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = "Microsoft Office is not installed or accessible."
                };
            }
            
            // Use circuit breaker to prevent cascading failures
            try
            {
                return await _circuitBreaker.ExecuteAsync(async () =>
                {
                    return await Task.Run(() =>
            {
                lock (_wordLock)
                {
                    var result = new ConversionResult();
                    dynamic doc = null;
                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                    
                    try
                    {
                        var wordApp = GetOrCreateWordApp();
                        
                        // Ensure Word is still healthy before conversion
                        EnsureWordHealthy();
                        
                        _logger.Debug("Opening document: {Path}", inputPath);
                        
                        doc = OpenDocumentSafely(wordApp, inputPath);
                        
                        _logger.Debug("Saving as PDF: {Path}", outputPath);
                        
                        // Save as PDF with version-safe method
                        if (!SaveAsPdfSafely(doc, outputPath))
                        {
                            throw new InvalidOperationException("Failed to save document as PDF");
                        }
                        
                        result.Success = true;
                        result.OutputPath = outputPath;
                        
                        stopwatch.Stop();
                        _logger.Information("Converted {File} in {ElapsedMs}ms using session Word instance", 
                            System.IO.Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to convert Word document");
                        result.Success = false;
                        result.ErrorMessage = ex.Message;
                        
                        // If Word crashed, clear the instance so it gets recreated
                        if (ex is COMException)
                        {
                            _logger.Warning("COM exception detected, will recreate Word instance on next use");
                            DisposeWordApp();
                        }
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try
                            {
                                doc.Close(SaveChanges: false);
                                ComHelper.SafeReleaseComObject(doc, "Document", "SessionAwareConversion");
                                doc = null;
                            }
                            catch (Exception closeEx)
                            {
                                _logger.Warning(closeEx, "Error closing document");
                            }
                        }
                    }
                    
                    return result;
                }
            });
                });
            }
            catch (InvalidOperationException circuitEx) when (circuitEx.Message.Contains("Circuit breaker is open"))
            {
                _logger.Warning("Circuit breaker prevented Word conversion: {Message}", circuitEx.Message);
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = "Word conversion service temporarily unavailable due to recent failures"
                };
            }
        }
        
        public void WarmUp()
        {
            lock (_wordLock)
            {
                if (_wordApp == null)
                {
                    GetOrCreateWordApp();
                    _logger.Information("Word pre-warmed for Save Quotes Mode");
                }
            }
        }
        
        private void DisposeWordApp()
        {
            if (_wordApp != null)
            {
                try
                {
                    // Try to close all open documents first
                    try
                    {
                        using (var comScope = new ComResourceScope())
                        {
                            var documents = comScope.GetDocuments(_wordApp, "SessionAwareDispose");
                            if (documents != null && documents.Count > 0)
                            {
                                _logger.Warning("Closing {Count} open documents", documents.Count);
                                // Use for loop instead of foreach to avoid COM enumerator leak
                                int count = documents.Count;
                                for (int i = count; i >= 1; i--) // Word collections are 1-based, iterate backwards
                                {
                                    try
                                    {
                                        dynamic doc = documents[i];
                                        doc.Close(SaveChanges: false);
                                        comScope.Track(doc, "Document", "SessionAwareDispose");
                                    }
                                    catch { }
                                }
                            }
                        } // documents collection automatically released here
                    }
                    catch { }
                    
                    _wordApp.Quit(SaveChanges: false);
                    ComHelper.SafeReleaseComObject(_wordApp, "WordApp", "SessionAwareDispose");
                    _wordApp = null;
                    
                    // Force garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    
                    _logger.Information("Word application disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing Word application");
                }
            }
        }
        
        public bool IsOfficeInstalled()
        {
            return IsOfficeAvailable();
        }
        
        public void Dispose()
        {
            if (!_disposed)
            {
                _idleTimer?.Dispose();
                
                lock (_wordLock)
                {
                    DisposeWordApp();
                }
                
                _disposed = true;
                _logger.Information("SessionAwareOfficeService disposed");
            }
        }
    }
} 