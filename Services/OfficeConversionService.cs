// Folder: Services/
// File: OfficeConversionService.cs
// Critical Fix #4: Proper COM object disposal to prevent memory leaks
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Serilog;
using Task = System.Threading.Tasks.Task;

namespace DocHandler.Services
{
    public class OfficeConversionService : IDisposable
    {
        private readonly ILogger _logger;
        private dynamic? _wordApp;
        private dynamic? _excelApp;
        private bool _disposed;
        private bool? _officeAvailable;
        private readonly object _wordLock = new object();
        private readonly object _excelLock = new object();
        
        private readonly ConfigurationService? _configService;
        private readonly ProcessManager? _processManager;
        
        public OfficeConversionService(ConfigurationService? configService = null, ProcessManager? processManager = null)
        {
            _logger = Log.ForContext<OfficeConversionService>();
            _configService = configService;
            _processManager = processManager;
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
        
        private bool IsOfficeAvailable()
        {
            if (_officeAvailable.HasValue)
                return _officeAvailable.Value;
                
            try
            {
                // Try to create Word application using late binding
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType != null)
                {
                    dynamic testApp = null;
                    try
                    {
                        testApp = Activator.CreateInstance(wordType);
                        ComHelper.TrackComObjectCreation("WordApp", "OfficeAvailabilityCheck");
                        testApp.Visible = false;
                        testApp.Quit();
                        _officeAvailable = true;
                        return true;
                    }
                    finally
                    {
                        if (testApp != null)
                        {
                            // CRITICAL FIX #4: Proper COM cleanup
                            ComHelper.SafeReleaseComObject(testApp, "WordApp", "OfficeAvailabilityCheck");
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

            var result = new ConversionResult();
            
            // Add file size validation
            var fileInfo = new FileInfo(inputPath);
            if (fileInfo.Length > 50 * 1024 * 1024) // 50MB limit
            {
                result.Success = false;
                result.ErrorMessage = "File size exceeds 50MB limit";
                return result;
            }

            _logger.Information("Converting Word to PDF: {WordPath} -> {PdfPath}", inputPath, outputPath);

            // CRITICAL FIX: Use proper cancellation token instead of Task.Wait() to prevent deadlocks
            var timeoutSeconds = _configService?.Config.ConversionTimeoutSeconds ?? 30;
            using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
            
            // CRITICAL FIX: Remove Task.Run - already on STA thread from caller
            lock (_wordLock)
            {
                // Validate apartment state
                if (!ValidateSTAThread("ConvertWordToPdf"))
                {
                    result.Success = false;
                    result.ErrorMessage = "Thread apartment state is not STA - Word operations will fail";
                    return result;
                }
                
                dynamic? doc = null;
                
                try
                {
                    // Check for cancellation before starting
                    cts.Token.ThrowIfCancellationRequested();
                    
                    // Ensure Word application is initialized using late binding
                    if (_wordApp == null)
                    {
                        try
                        {
                            Type? wordType = Type.GetTypeFromProgID("Word.Application");
                            if (wordType == null)
                            {
                                result.Success = false;
                                result.ErrorMessage = "Microsoft Word is not installed.";
                                return result;
                            }
                            
                            _wordApp = Activator.CreateInstance(wordType);
                            ComHelper.TrackComObjectCreation("WordApp", "OfficeConversionService");
                            _wordApp.Visible = false;
                            _wordApp.DisplayAlerts = 0; // wdAlertsNone
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Failed to create Word application");
                            result.Success = false;
                            result.ErrorMessage = "Microsoft Word is not installed or accessible.";
                            return result;
                        }
                    }
                    
                    // Check for cancellation before document operations
                    cts.Token.ThrowIfCancellationRequested();
                    
                    // CRITICAL FIX: Use ComResourceScope for automatic COM object cleanup
                    _logger.Debug("Opening Word document: {Path}", inputPath);
                    
                    using (var comScope = new ComResourceScope())
                    {
                        var documents = comScope.Track(_wordApp.Documents, "Documents", "ConvertWordToPdf");
                        doc = comScope.Track(
                            documents.Open(
                                inputPath,
                                ReadOnly: true,
                                AddToRecentFiles: false,
                                Repair: false,
                                ShowRepairs: false,
                                OpenAndRepair: false,
                                NoEncodingDialog: true,
                                Revert: false
                            ),
                            "Document",
                            "ConvertWordToPdf"
                        );
                    } // Documents collection automatically released here
                    
                    // Check for cancellation before PDF conversion
                    cts.Token.ThrowIfCancellationRequested();
                    
                    _logger.Debug("Converting to PDF: {Path}", outputPath);
                    doc.SaveAs2(outputPath, FileFormat: 17);
                    
                    result.Success = true;
                    result.OutputPath = outputPath;
                    _logger.Information("Successfully converted Word to PDF");
                }
                catch (OperationCanceledException)
                {
                    _logger.Error("Word conversion timed out after {TimeoutSeconds} seconds", timeoutSeconds);
                    result.Success = false;
                    result.ErrorMessage = $"Word conversion timed out after {timeoutSeconds} seconds";
                    
                    // Force Word restart on timeout - dispose current instance
                    if (_wordApp != null)
                    {
                        try
                        {
                            _wordApp.Quit();
                            ComHelper.SafeReleaseComObject(_wordApp, "WordApp", "TimeoutRestart");
                            _wordApp = null;
                            _logger.Information("Word application restarted due to timeout");
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error disposing Word app after timeout");
                        }
                        
                        // Use ProcessManager to clean up any orphaned Word processes
                        try
                        {
                            _processManager?.KillOrphanedWordProcesses();
                            _logger.Information("Cleaned up orphaned Word processes after timeout");
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error cleaning up orphaned processes after timeout");
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to convert Word document to PDF");
                    result.Success = false;
                    result.ErrorMessage = $"Conversion failed: {ex.Message}";
                }
                finally
                {
                    // Always cleanup document
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                            ComHelper.SafeReleaseComObject(doc, "Document", "ConvertWordToPdf");
                            doc = null;
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word document");
                        }
                    }
                    
                    // Force garbage collection after COM operations
                    ComHelper.ForceComCleanup("ConvertWordToPdf");
                }
            }
            
            return result;
        }
        
        private ConversionResult ConvertWordToPdfSync(string inputPath, string outputPath)
        {
            var result = new ConversionResult();
            dynamic? doc = null;
            
            // Validate apartment state for synchronous operations
            if (!ValidateSTAThread("ConvertWordToPdfSync"))
            {
                result.Success = false;
                result.ErrorMessage = "Thread apartment state is not STA - COM operations may fail";
                return result;
            }
            
            // CRITICAL FIX #4: Thread-safe lock for Word operations
            lock (_wordLock)
            {
                try
                {
                    // Ensure Word application is initialized using late binding
                    if (_wordApp == null)
                    {
                        try
                        {
                            Type? wordType = Type.GetTypeFromProgID("Word.Application");
                            if (wordType == null)
                            {
                                result.Success = false;
                                result.ErrorMessage = "Microsoft Word is not installed.";
                                return result;
                            }
                            
                            _wordApp = Activator.CreateInstance(wordType);
                            ComHelper.TrackComObjectCreation("WordApp", "ConvertWordToPdfSync");
                            _wordApp.Visible = false;
                            _wordApp.DisplayAlerts = 0; // wdAlertsNone
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Failed to create Word application");
                            result.Success = false;
                            result.ErrorMessage = "Microsoft Word is not installed or accessible.";
                            return result;
                        }
                    }
                    
                    // Open the document using ComResourceScope for automatic cleanup
                    _logger.Information("Opening Word document: {Path}", inputPath);
                    using (var comScope = new ComResourceScope())
                    {
                        var documents = comScope.Track(_wordApp.Documents, "Documents", "ConvertWordToPdfSync");
                        doc = comScope.Track(
                            documents.Open(
                                inputPath,
                                ReadOnly: true,
                                AddToRecentFiles: false,
                                Repair: false,
                                ShowRepairs: false,
                                OpenAndRepair: false,
                                NoEncodingDialog: true,
                                Revert: false
                            ),
                            "Document",
                            "ConvertWordToPdfSync"
                        );
                    } // Documents collection automatically released here
                    
                    // Save as PDF (17 = wdFormatPDF)
                    _logger.Information("Converting to PDF: {Path}", outputPath);
                    doc.SaveAs2(outputPath, FileFormat: 17);
                    
                    result.Success = true;
                    result.OutputPath = outputPath;
                    _logger.Information("Successfully converted Word to PDF");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to convert Word document to PDF");
                    result.Success = false;
                    result.ErrorMessage = $"Conversion failed: {ex.Message}";
                }
                finally
                {
                    // CRITICAL FIX #4: Proper cleanup in finally block
                    // Clean up document
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                            ComHelper.SafeReleaseComObject(doc, "Document", "ConvertWordToPdfSync");
                            doc = null;
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word document");
                        }
                    }
                    
                    // CRITICAL FIX #4: Force garbage collection after COM operations
                    ComHelper.ForceComCleanup("ConvertWordToPdfSync");
                }
            }
            
            return result;
        }
        
        public async System.Threading.Tasks.Task<ConversionResult> ConvertExcelToPdf(string inputPath, string outputPath)
        {
            if (!IsOfficeAvailable())
            {
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = "Microsoft Office is not installed or accessible. Please install Microsoft Office to convert Excel documents to PDF."
                };
            }
            
            _logger.Information("Converting Excel to PDF: {ExcelPath} -> {PdfPath}", inputPath, outputPath);

            // CRITICAL FIX: Remove Task.Run - already on STA thread from caller
            var result = new ConversionResult();
            dynamic? workbook = null;

            lock (_excelLock)
            {
                // Validate apartment state
                if (!ValidateSTAThread("ConvertExcelToPdf"))
                {
                    result.Success = false;
                    result.ErrorMessage = "Thread apartment state is not STA - Excel operations will fail";
                    return result;
                }
                
                try
                {
                    // Create Excel application if needed using late binding
                    if (_excelApp == null)
                    {
                        try
                        {
                            Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                            if (excelType == null)
                            {
                                result.Success = false;
                                result.ErrorMessage = "Microsoft Excel is not installed.";
                                return result;
                            }
                            
                            _excelApp = Activator.CreateInstance(excelType);
                            ComHelper.TrackComObjectCreation("ExcelApp", "OfficeConversionService");
                            _excelApp.Visible = false;
                            _excelApp.DisplayAlerts = false;
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Failed to create Excel application");
                            result.Success = false;
                            result.ErrorMessage = "Microsoft Excel is not installed or accessible.";
                            return result;
                        }
                    }

                    // Open the workbook using ComResourceScope for automatic cleanup
                    using (var comScope = new ComResourceScope())
                    {
                        var workbooks = comScope.Track(_excelApp.Workbooks, "Workbooks", "ConvertExcelToPdf");
                        workbook = comScope.Track(
                            workbooks.Open(
                                inputPath,
                                ReadOnly: true,
                                UpdateLinks: false,
                                Notify: false
                            ),
                            "Workbook",
                            "ConvertExcelToPdf"
                        );
                    } // Workbooks collection automatically released here

                    // Export as PDF (0 = xlTypePDF)
                    workbook.ExportAsFixedFormat(
                        Type: 0,
                        Filename: outputPath,
                        Quality: 0, // xlQualityStandard
                        IncludeDocProperties: true,
                        IgnorePrintAreas: false,
                        OpenAfterPublish: false);

                    _logger.Information("Successfully converted Excel to PDF");
                    result.Success = true;
                    result.OutputPath = outputPath;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to convert Excel to PDF");
                    result.Success = false;
                    result.ErrorMessage = $"Excel conversion failed: {ex.Message}";
                }
                finally
                {
                    // CRITICAL FIX #4: Proper cleanup in finally block
                    // Clean up
                    if (workbook != null)
                    {
                        try
                        {
                            workbook.Close(false);
                            ComHelper.SafeReleaseComObject(workbook, "Workbook", "ConvertExcelToPdf");
                            workbook = null;
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Excel workbook");
                        }
                    }
                    
                    // CRITICAL FIX #4: Force garbage collection after COM operations
                    ComHelper.ForceComCleanup("ConvertExcelToPdf");
                }
            }
            
            return result;
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
        
        // CRITICAL FIX #4: Comprehensive disposal pattern
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources
                }
                
                // Clean up Word application
                if (_wordApp != null)
                {
                    lock (_wordLock)
                    {
                        try
                        {
                            // Try to close all documents first
                            try
                            {
                                var documents = _wordApp.Documents;
                                if (documents != null && documents.Count > 0)
                                {
                                    // Use for loop instead of foreach to avoid COM enumerator leak
                                    int count = documents.Count;
                                    for (int i = count; i >= 1; i--) // Word collections are 1-based, iterate backwards
                                    {
                                        try
                                        {
                                            dynamic doc = documents[i];
                                            doc.Close(SaveChanges: false);
                                            ComHelper.SafeReleaseComObject(doc, "Document", "Dispose");
                                        }
                                        catch { }
                                    }
                                    ComHelper.SafeReleaseComObject(documents, "Documents", "Dispose");
                                }
                            }
                            catch { }
                            
                            _wordApp.Quit();
                            ComHelper.SafeReleaseComObject(_wordApp, "WordApp", "Dispose");
                            _wordApp = null;
                            
                            _logger.Information("Word application closed");
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word application");
                        }
                    }
                }
                
                // Clean up Excel application
                if (_excelApp != null)
                {
                    lock (_excelLock)
                    {
                        try
                        {
                            // Try to close all workbooks first
                            try
                            {
                                var workbooks = _excelApp.Workbooks;
                                if (workbooks != null && workbooks.Count > 0)
                                {
                                    // Use for loop instead of foreach to avoid COM enumerator leak
                                    int count = workbooks.Count;
                                    for (int i = count; i >= 1; i--) // Excel collections are 1-based, iterate backwards
                                    {
                                        try
                                        {
                                            dynamic wb = workbooks[i];
                                            wb.Close(false);
                                            ComHelper.SafeReleaseComObject(wb, "Workbook", "Dispose");
                                        }
                                        catch { }
                                    }
                                    ComHelper.SafeReleaseComObject(workbooks, "Workbooks", "Dispose");
                                }
                            }
                            catch { }
                            
                            _excelApp.Quit();
                            ComHelper.SafeReleaseComObject(_excelApp, "ExcelApp", "Dispose");
                            _excelApp = null;
                            
                            _logger.Information("Excel application closed");
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Excel application");
                        }
                    }
                }
                
                // CRITICAL FIX #4: Triple garbage collection to ensure COM cleanup
                // Force garbage collection to release COM objects
                ComHelper.ForceComCleanup("Dispose");
                
                // Log final COM object statistics
                ComHelper.LogComObjectStats();
                
                _disposed = true;
            }
        }
        
        // CRITICAL FIX #4: Finalizer for unmanaged resource cleanup
        ~OfficeConversionService()
        {
            Dispose(false);
        }
    }
    
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string? OutputPath { get; set; }
        public string? ErrorMessage { get; set; }
    }
}