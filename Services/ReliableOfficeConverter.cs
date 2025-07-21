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
        public ConversionResult ConvertWordToPdf(string inputPath, string outputPath, bool singleUse = false)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ReliableOfficeConverter));

            lock (_lock)
            {
                var result = new ConversionResult();
                dynamic doc = null;

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

                    // Open and convert document
                    doc = _wordApp.Documents.Open(
                        inputPath,
                        ReadOnly: true,
                        AddToRecentFiles: false,
                        Visible: false
                    );

                    // Save as PDF (17 = wdFormatPDF)
                    doc.SaveAs2(outputPath, 17);

                    result.Success = true;
                    result.OutputPath = outputPath;

                    _wordUseCount++;
                    stopwatch.Stop();

                    _logger.Information("Converted {File} to PDF in {Ms}ms (Use #{Count})",
                        Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds, _wordUseCount);

                    if (singleUse)
                    {
                        _logger.Debug("Single-use mode: cleaning up Word instance");
                        CleanupWord();
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

                    return result;
                }
                finally
                {
                    // Always close document
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                            Marshal.ReleaseComObject(doc);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing document");
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Convert Excel spreadsheet to PDF with optional single-use mode
        /// </summary>
        public ConversionResult ConvertExcelToPdf(string inputPath, string outputPath, bool singleUse = false)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(ReliableOfficeConverter));

            lock (_lock)
            {
                var result = new ConversionResult();
                dynamic workbook = null;

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

                    // Open and convert workbook
                    workbook = _excelApp.Workbooks.Open(
                        inputPath,
                        ReadOnly: true,
                        UpdateLinks: false,
                        Notify: false
                    );

                    // Export as PDF (0 = xlTypePDF)
                    workbook.ExportAsFixedFormat(
                        Type: 0,
                        Filename: outputPath,
                        Quality: 0, // Standard quality
                        IncludeDocProperties: false,
                        IgnorePrintAreas: false
                    );

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
                finally
                {
                    // Always close workbook
                    if (workbook != null)
                    {
                        try
                        {
                            workbook.Close(SaveChanges: false);
                            Marshal.ReleaseComObject(workbook);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing workbook");
                        }
                    }
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
                _logger.Information("Creating new Word instance");

                Type wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    throw new InvalidOperationException("Word.Application not available");
                }

                _wordApp = Activator.CreateInstance(wordType);
                _wordApp.Visible = false;
                _wordApp.DisplayAlerts = 0; // wdAlertsNone

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

                    // Close all documents
                    try
                    {
                        while (_wordApp.Documents.Count > 0)
                        {
                            _wordApp.Documents[1].Close(SaveChanges: false);
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

                    // Close all workbooks
                    try
                    {
                        while (_excelApp.Workbooks.Count > 0)
                        {
                            _excelApp.Workbooks[1].Close(SaveChanges: false);
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