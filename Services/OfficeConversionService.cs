// Folder: Services/
// File: OfficeConversionService.cs
// Critical Fix #4: Proper COM object disposal to prevent memory leaks
using System;
using System.IO;
using System.Runtime.InteropServices;
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
        
        public OfficeConversionService()
        {
            _logger = Log.ForContext<OfficeConversionService>();
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
                                // CRITICAL FIX #4: Proper COM cleanup
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
        
        public async System.Threading.Tasks.Task<ConversionResult> ConvertWordToPdf(string inputPath, string outputPath)
        {
            if (!IsOfficeAvailable())
            {
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = "Microsoft Office is not installed or accessible. Please install Microsoft Office to convert Word documents to PDF."
                };
            }
            
            return await Task.Run(() => ConvertWordToPdfSync(inputPath, outputPath));
        }
        
        private ConversionResult ConvertWordToPdfSync(string inputPath, string outputPath)
        {
            var result = new ConversionResult();
            dynamic? doc = null;
            
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
                    
                    // Open the document
                    _logger.Information("Opening Word document: {Path}", inputPath);
                    doc = _wordApp.Documents.Open(inputPath, ReadOnly: true);
                    
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
                            Marshal.FinalReleaseComObject(doc);
                            doc = null;
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word document");
                        }
                    }
                    
                    // CRITICAL FIX #4: Force garbage collection after COM operations
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
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

            return await Task.Run(() =>
            {
                dynamic? workbook = null;
                var result = new ConversionResult();

                // CRITICAL FIX #4: Thread-safe lock for Excel operations
                lock (_excelLock)
                {
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

                        // Open the workbook
                        workbook = _excelApp.Workbooks.Open(
                            inputPath,
                            ReadOnly: true,
                            IgnoreReadOnlyRecommended: true,
                            Notify: false);

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
                                Marshal.FinalReleaseComObject(workbook);
                                workbook = null;
                            }
                            catch (Exception ex)
                            {
                                _logger.Warning(ex, "Error closing Excel workbook");
                            }
                        }
                        
                        // CRITICAL FIX #4: Force garbage collection after COM operations
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }

                return result;
            });
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
                                    foreach (dynamic doc in documents)
                                    {
                                        try
                                        {
                                            doc.Close(SaveChanges: false);
                                            Marshal.FinalReleaseComObject(doc);
                                        }
                                        catch { }
                                    }
                                    Marshal.FinalReleaseComObject(documents);
                                }
                            }
                            catch { }
                            
                            _wordApp.Quit();
                            Marshal.FinalReleaseComObject(_wordApp);
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
                                    foreach (dynamic wb in workbooks)
                                    {
                                        try
                                        {
                                            wb.Close(false);
                                            Marshal.FinalReleaseComObject(wb);
                                        }
                                        catch { }
                                    }
                                    Marshal.FinalReleaseComObject(workbooks);
                                }
                            }
                            catch { }
                            
                            _excelApp.Quit();
                            Marshal.FinalReleaseComObject(_excelApp);
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
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
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