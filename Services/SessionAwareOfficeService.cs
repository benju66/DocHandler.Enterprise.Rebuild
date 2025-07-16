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
        
        public SessionAwareOfficeService()
        {
            _logger = Log.ForContext<SessionAwareOfficeService>();
            _circuitBreaker = new ConversionCircuitBreaker();
            _logger.Information("Initializing session-aware Office service with circuit breaker protection");
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
                        _logger.Information("Creating new Word application instance for session");
                        Type wordType = Type.GetTypeFromProgID("Word.Application");
                        if (wordType == null)
                        {
                            throw new InvalidOperationException("Word.Application ProgID not found");
                        }
                        
                        _wordApp = Activator.CreateInstance(wordType);
                        
                        // Make Word invisible and optimize for speed
                        _wordApp.Visible = false;
                        _wordApp.DisplayAlerts = 0; // wdAlertsNone
                        
                        // SAFE: Only set properties that don't require active documents
                        try
                        {
                            _wordApp.ScreenUpdating = false;
                        }
                        catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800A11FD))
                        {
                            // Ignore "document window is not active" error during pre-warm
                            _logger.Debug("Skipping ScreenUpdating property - no active document during pre-warm");
                        }
                        
                        try
                        {
                            _wordApp.DisplayRecentFiles = false;
                            _wordApp.DisplayScrollBars = false;
                            _wordApp.DisplayStatusBar = false;
                        }
                        catch (System.Runtime.InteropServices.COMException ex) when (ex.HResult == unchecked((int)0x800A11FD))
                        {
                            // Ignore properties that require active document
                            _logger.Debug("Skipping UI properties - no active document during pre-warm");
                        }
                        
                        // SAFE: Window state can be set without active document
                        try
                        {
                            // Minimize window to prevent any flashing
                            _wordApp.WindowState = -2; // wdWindowStateMinimize
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            // Ignore if window state cannot be set
                            _logger.Debug("Could not set WindowState during pre-warm");
                        }
                        
                        // Set last health check time
                        _lastHealthCheck = DateTime.Now;
                        
                        // REDUCED: Only essential optimization settings to prevent instability
                        try
                        {
                            // Critical performance optimizations (safe)
                            _wordApp.Options.CheckGrammarAsYouType = false;
                            _wordApp.Options.CheckSpellingAsYouType = false;
                            _wordApp.Options.BackgroundSave = false;
                            _wordApp.Options.SaveInterval = 0; // Disable auto-save
                            
                            // Removed potentially unstable options:
                            // - EnableAutoRecovery, AllowFastSave, CreateBackup (can cause file locks)
                            // - UpdateLinksAtOpen (can cause hanging on network resources)
                            // - AnimateScreenMovements, PaginationView (UI-related, can cause issues)
                            // - SavePropertiesPrompt, WPHelp (dialog-related, can cause hangs)
                        }
                        catch (Exception optEx)
                        {
                            _logger.Warning(optEx, "Some Word optimization options could not be set");
                        }
                        
                        // Get process ID for monitoring
                        try
                        {
                            var processId = (int)_wordApp.GetType().InvokeMember("ProcessID", 
                                System.Reflection.BindingFlags.GetProperty, null, _wordApp, null);
                            _logger.Information("Word application created with PID: {ProcessId}", processId);
                        }
                        catch { }
                        
                        // Start idle cleanup timer
                        _idleTimer = new Timer(CheckIdleTimeout, null, TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
                        
                        _logger.Information("Word application created and optimized for session");
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
                
                // Try to access a property to verify Word is responsive
                var _ = _wordApp.Version;
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
                        
                        doc = wordApp.Documents.Open(
                            inputPath,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            Repair: false,
                            ShowRepairs: false,
                            OpenAndRepair: false,
                            NoEncodingDialog: true,
                            Revert: false
                        );
                        
                        _logger.Debug("Saving as PDF: {Path}", outputPath);
                        
                        // Save as PDF with optimized settings
                        doc.SaveAs2(
                            outputPath, 
                            FileFormat: 17, // wdFormatPDF
                            EmbedTrueTypeFonts: false,
                            SaveNativePictureFormat: false,
                            SaveFormsData: false,
                            CompressLevel: 0
                        );
                        
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
                        var documents = _wordApp.Documents;
                        if (documents != null && documents.Count > 0)
                        {
                            _logger.Warning("Closing {Count} open documents", documents.Count);
                            foreach (dynamic doc in documents)
                            {
                                try
                                {
                                    doc.Close(SaveChanges: false);
                                    ComHelper.SafeReleaseComObject(doc, "Document", "SessionAwareDispose");
                                }
                                catch { }
                            }
                            ComHelper.SafeReleaseComObject(documents, "Documents", "SessionAwareDispose");
                        }
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