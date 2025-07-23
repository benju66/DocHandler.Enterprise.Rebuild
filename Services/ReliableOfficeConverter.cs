using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Reliable Office conversion service with batch support and automatic cleanup.
    /// Replaces the flawed session-aware pattern with proper lifecycle management.
    /// </summary>
    public class ReliableOfficeConverter : IDisposable
    {
        private readonly ILogger _logger;
        private readonly object _lock = new object();
        private dynamic _wordApp;
        private dynamic _excelApp;
        private int _wordUseCount = 0;
        private int _excelUseCount = 0;
        private readonly int _maxUses = 20; // Recreate after 20 uses to prevent memory growth
        private readonly OfficeProcessGuard _processGuard;
        private bool _disposed;

        public ReliableOfficeConverter()
        {
            _logger = Log.ForContext<ReliableOfficeConverter>();
            _processGuard = new OfficeProcessGuard();
        }

        /// <summary>
        /// Convert Word document to PDF with optional single-use mode
        /// </summary>
        public OfficeConversionResult ConvertWordToPdf(string inputPath, string outputPath, bool singleUse = false)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ReliableOfficeConverter));

            lock (_lock)
            {
                var result = new OfficeConversionResult();

                try
                {
                    // Validate STA thread
                    if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    {
                        throw new InvalidOperationException("Must be called from STA thread");
                    }

                    // Create or recreate Word instance if needed
                    if (_wordApp == null || _wordUseCount >= _maxUses)
                    {
                        CleanupWord();
                        CreateWordInstance();
                    }

                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                    // CRITICAL FIX: Use ComResourceScope for automatic COM cleanup
                    using (var comScope = new ComResourceScope())
                    {
                        // Ensure screen updating is off during conversion
                        var screenUpdatingWasEnabled = false;
                        try
                        {
                            screenUpdatingWasEnabled = _wordApp.ScreenUpdating;
                            _wordApp.ScreenUpdating = false;
                        }
                        catch { }

                        // This automatically tracks and releases the Documents collection
                        var documents = comScope.Track(_wordApp.Documents, "Documents", "ConvertWordToPdf");
                        
                        // Open and convert document with additional parameters to prevent UI
                        var doc = comScope.Track(
                            documents.Open(
                                inputPath,
                                ReadOnly: true,
                                AddToRecentFiles: false,
                                Visible: false,
                                OpenAndRepair: false,
                                NoEncodingDialog: true
                            ),
                            "Document",
                            "ConvertWordToPdf"
                        );

                        // Save as PDF (17 = wdFormatPDF)
                        doc.SaveAs2(outputPath, 17);
                        
                        // Close document before scope disposal
                        doc.Close(SaveChanges: false);
                        
                        // Restore screen updating if it was changed
                        try
                        {
                            if (screenUpdatingWasEnabled)
                                _wordApp.ScreenUpdating = true;
                        }
                        catch { }
                    } // ComResourceScope automatically releases all tracked COM objects here

                    result.Success = true;
                    result.OutputPath = outputPath;

                    _wordUseCount++;
                    stopwatch.Stop();

                    // Only log if conversion was slow or verbose logging is enabled
                    var slowThreshold = 1000; // milliseconds
                    if (stopwatch.ElapsedMilliseconds > slowThreshold)
                    {
                        _logger.Warning("Slow conversion: {File} to PDF in {Ms}ms (Use #{Count})",
                            Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds, _wordUseCount);
                    }
                    else
                    {
                        _logger.Debug("Converted {File} to PDF in {Ms}ms (Use #{Count})",
                            Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds, _wordUseCount);
                    }

                    if (singleUse)
                    {
                        _logger.Debug("Single-use mode: cleaning up Word instance");
                        CleanupWord();
                    }

                    return result;
                }
                catch (COMException comEx)
                {
                    _logger.Error(comEx, "Word conversion failed with COM error for {File}, HResult: {HResult:X8}", 
                        inputPath, comEx.HResult);
                    
                    // Create custom Office exception with recovery guidance
                    var officeEx = OfficeOperationException.FromCOMException("Word", "Convert to PDF", comEx);
                    
                    result.Success = false;
                    result.ErrorMessage = officeEx.UserFriendlyMessage;

                    // Always cleanup on error
                    CleanupWord();

                    // For single-use mode or non-recoverable errors, throw custom exception
                    if (singleUse || !officeEx.IsRecoverable)
                    {
                        throw officeEx;
                    }

                    return result;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Word conversion failed for {File}", inputPath);
                    result.Success = false;
                    result.ErrorMessage = ex.Message;

                    // Always cleanup on error
                    CleanupWord();

                    // For single-use mode, throw file processing exception
                    if (singleUse)
                    {
                        throw new FileProcessingException(inputPath, "Word to PDF conversion", 
                            "Document conversion failed", 
                            "Try again with a different document or ensure Microsoft Word is working properly.", 
                            ex);
                    }

                    return result;
                }
            }
        }

        /// <summary>
        /// Convert Excel spreadsheet to PDF with optional single-use mode
        /// </summary>
        public OfficeConversionResult ConvertExcelToPdf(string inputPath, string outputPath, bool singleUse = false)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ReliableOfficeConverter));

            lock (_lock)
            {
                var result = new OfficeConversionResult();

                try
                {
                    // Validate STA thread
                    if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    {
                        throw new InvalidOperationException("Must be called from STA thread");
                    }

                    // Create or recreate Excel instance if needed
                    if (_excelApp == null || _excelUseCount >= _maxUses)
                    {
                        CleanupExcel();
                        CreateExcelInstance();
                    }

                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();

                    // CRITICAL FIX: Use ComResourceScope for automatic COM cleanup
                    using (var comScope = new ComResourceScope())
                    {
                        // This automatically tracks and releases the Workbooks collection
                        var workbooks = comScope.Track(_excelApp.Workbooks, "Workbooks", "ConvertExcelToPdf");
                        
                        // Open workbook
                        var workbook = comScope.Track(
                            workbooks.Open(
                                inputPath,
                                ReadOnly: true,
                                UpdateLinks: false,
                                Notify: false
                            ),
                            "Workbook",
                            "ConvertExcelToPdf"
                        );

                        // Export as PDF (0 = xlTypePDF)
                        workbook.ExportAsFixedFormat(
                            Type: 0,
                            Filename: outputPath,
                            Quality: 0, // Standard quality
                            IncludeDocProperties: false,
                            IgnorePrintAreas: false
                        );
                        
                        // Close workbook before scope disposal
                        workbook.Close(SaveChanges: false);
                    } // ComResourceScope automatically releases all tracked COM objects here

                    result.Success = true;
                    result.OutputPath = outputPath;

                    _excelUseCount++;
                    stopwatch.Stop();

                    _logger.Information("Converted {File} to PDF in {Ms}ms (Use #{Count})",
                        Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds, _excelUseCount);

                    if (singleUse)
                    {
                        _logger.Debug("Single-use mode: cleaning up Excel instance");
                        CleanupExcel();
                    }

                    return result;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Excel conversion failed for {File}", inputPath);
                    result.Success = false;
                    result.ErrorMessage = ex.Message;

                    // Always cleanup on error
                    CleanupExcel();

                    return result;
                }
            }
        }

        /// <summary>
        /// Finish a batch of conversions and cleanup instances
        /// </summary>
        public void FinishBatch()
        {
            lock (_lock)
            {
                _logger.Information("Finishing batch - cleaning up Office instances");
                CleanupWord();
                CleanupExcel();
            }
        }

        private void CreateWordInstance()
        {
            try
            {
                _logger.Debug("Creating new Word instance");

                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    throw new InvalidOperationException("Word.Application not available");
                }

                _wordApp = Activator.CreateInstance(wordType);
                _wordApp.Visible = false;
                _wordApp.DisplayAlerts = 0; // wdAlertsNone
                
                // Additional optimizations to prevent UI interference
                try
                {
                    // Disable screen updating to prevent spinning cursor
                    _wordApp.ScreenUpdating = false;
                    
                    // Apply optimizations to Word Options if available
                    if (_wordApp.Options != null)
                    {
                        _wordApp.Options.CheckGrammarAsYouType = false;
                        _wordApp.Options.CheckSpellingAsYouType = false;
                        _wordApp.Options.BackgroundSave = false;
                        _wordApp.Options.SaveInterval = 0; // Disable auto-save
                        _wordApp.Options.AnimateScreenMovements = false;
                        _wordApp.Options.ConfirmConversions = false;
                        _wordApp.Options.UpdateFieldsAtPrint = false;
                        _wordApp.Options.UpdateLinksAtPrint = false;
                    }
                }
                catch (Exception optEx)
                {
                    _logger.Debug("Some Word optimization properties not available: {Message}", optEx.Message);
                    // Continue anyway - these are just optimizations
                }

                // Track the process
                _wordUseCount = 0;
                ComHelper.TrackComObjectCreation("WordApp", "ReliableOfficeConverter");
                
                // Try to register the process (with retry for Hwnd availability)
                try
                {
                    // Wait a moment for Word to fully initialize
                    System.Threading.Thread.Sleep(100);
                    
                    int retries = 3;
                    uint processId = 0;
                    
                    while (retries > 0 && processId == 0)
                    {
                        try
                        {
                            IntPtr hwnd = new IntPtr(_wordApp.Hwnd);
                            GetWindowThreadProcessId(hwnd, out processId);
                            if (processId > 0) break;
                        }
                        catch
                        {
                            if (--retries > 0)
                                System.Threading.Thread.Sleep(200);
                        }
                    }
                    
                    if (processId > 0)
                    {
                        _processGuard.RegisterProcess((int)processId);
                        _logger.Debug("Word process registered: {PID}", processId);
                    }
                    else
                    {
                        _logger.Warning("Could not register Word process - Hwnd not available");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Could not register Word process");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to create Word instance");
                throw;
            }
        }

        private void CreateExcelInstance()
        {
            try
            {
                _logger.Information("Creating new Excel instance");

                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel.Application not available");
                }

                _excelApp = Activator.CreateInstance(excelType);
                _excelApp.Visible = false;
                _excelApp.DisplayAlerts = false;
                _excelApp.ScreenUpdating = false;

                // Track the process
                try
                {
                    IntPtr hwnd = new IntPtr((int)_excelApp.Hwnd);
                    uint processId;
                    GetWindowThreadProcessId(hwnd, out processId);
                    _processGuard.RegisterProcess((int)processId);
                    _logger.Debug("Excel process registered: {PID}", processId);
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Could not register Excel process");
                }

                _excelUseCount = 0;
                ComHelper.TrackComObjectCreation("ExcelApp", "ReliableOfficeConverter");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to create Excel instance");
                throw;
            }
        }

        private void CleanupWord()
        {
            _logger.Information("CleanupWord called - _wordApp is {Status}", _wordApp != null ? "not null" : "null");
            if (_wordApp != null)
            {
                try
                {
                    _logger.Information("Cleaning up Word instance (used {Count} times)", _wordUseCount);

                    // Close all documents with proper COM cleanup
                    try
                    {
                        using (var comScope = new ComResourceScope())
                        {
                            var documents = comScope.Track(_wordApp.Documents, "Documents", "CleanupWord");
                            while (documents.Count > 0)
                            {
                                var doc = comScope.Track(documents[1], "Document", "CleanupWord");
                                doc.Close(SaveChanges: false);
                            }
                        }
                    }
                    catch { }

                    // Quit Word
                    _wordApp.Quit(SaveChanges: false);
                    
                    // Wait a moment for Word to process the quit command
                    System.Threading.Thread.Sleep(500);
                    
                    ComHelper.SafeReleaseComObject(_wordApp, "WordApp", "ReliableOfficeConverter");
                    _wordApp = null;
                    _wordUseCount = 0;

                    // Force garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    _logger.Information("Word cleanup completed");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error during Word cleanup");
                }
            }
        }

        private void CleanupExcel()
        {
            if (_excelApp != null)
            {
                try
                {
                    _logger.Information("Cleaning up Excel instance (used {Count} times)", _excelUseCount);

                    // Close all workbooks with proper COM cleanup
                    try
                    {
                        using (var comScope = new ComResourceScope())
                        {
                            var workbooks = comScope.Track(_excelApp.Workbooks, "Workbooks", "CleanupExcel");
                            while (workbooks.Count > 0)
                            {
                                var workbook = comScope.Track(workbooks[1], "Workbook", "CleanupExcel");
                                workbook.Close(SaveChanges: false);
                            }
                        }
                    }
                    catch { }

                    // Quit Excel
                    _excelApp.Quit();
                    ComHelper.SafeReleaseComObject(_excelApp, "ExcelApp", "ReliableOfficeConverter");
                    _excelApp = null;
                    _excelUseCount = 0;

                    // Force garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();

                    _logger.Information("Excel cleanup completed");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error during Excel cleanup");
                }
            }
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
                    // Dispose managed resources
                    _processGuard?.Dispose();
                }

                // Cleanup COM objects
                CleanupWord();
                CleanupExcel();

                _disposed = true;
                _logger.Information("ReliableOfficeConverter disposed");
            }
        }

        // Windows API for getting process ID
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
} 