using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Serilog;
using Task = System.Threading.Tasks.Task;

namespace DocHandler.Services
{
    public class OfficeConversionService : IDisposable
    {
        private readonly ILogger _logger;
        private Application? _wordApp;
        private bool _disposed;
        
        public OfficeConversionService()
        {
            _logger = Log.ForContext<OfficeConversionService>();
        }
        
        public async System.Threading.Tasks.Task<ConversionResult> ConvertWordToPdf(string inputPath, string outputPath)
        {
            return await Task.Run(() => ConvertWordToPdfSync(inputPath, outputPath));
        }
        
        private ConversionResult ConvertWordToPdfSync(string inputPath, string outputPath)
        {
            var result = new ConversionResult();
            Document? doc = null;
            
            try
            {
                // Ensure Word application is initialized
                if (_wordApp == null)
                {
                    try
                    {
                        _wordApp = new Application();
                        _wordApp.Visible = false;
                        _wordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
                    }
                    catch (COMException ex)
                    {
                        _logger.Error(ex, "Microsoft Word is not installed or accessible");
                        result.Success = false;
                        result.ErrorMessage = "Microsoft Word is not installed. Please install Microsoft Office to convert Word documents to PDF.";
                        return result;
                    }
                }
                
                // Open the document
                _logger.Information("Opening Word document: {Path}", inputPath);
                doc = _wordApp.Documents.Open(inputPath, ReadOnly: true);
                
                // Save as PDF
                _logger.Information("Converting to PDF: {Path}", outputPath);
                doc.SaveAs2(outputPath, WdSaveFormat.wdFormatPDF);
                
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
                // Clean up document
                if (doc != null)
                {
                    try
                    {
                        doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                        Marshal.ReleaseComObject(doc);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error closing Word document");
                    }
                }
            }
            
            return result;
        }
        
        public async System.Threading.Tasks.Task<ConversionResult> ConvertExcelToPdf(string inputPath, string outputPath)
        {
            // Placeholder for Excel conversion
            // Will need Microsoft.Office.Interop.Excel package
            var result = new ConversionResult
            {
                Success = false,
                ErrorMessage = "Excel conversion not yet implemented"
            };
            
            return await Task.FromResult(result);
        }
        
        public bool IsOfficeInstalled()
        {
            try
            {
                var testApp = new Application();
                testApp.Quit();
                Marshal.ReleaseComObject(testApp);
                return true;
            }
            catch
            {
                return false;
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
                }
                
                // Clean up Word application
                if (_wordApp != null)
                {
                    try
                    {
                        _wordApp.Quit();
                        Marshal.ReleaseComObject(_wordApp);
                        _wordApp = null;
                        
                        // Force garbage collection to release COM objects
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        
                        _logger.Information("Word application closed");
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error closing Word application");
                    }
                }
                
                _disposed = true;
            }
        }
    }
    
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string? OutputPath { get; set; }
        public string? ErrorMessage { get; set; }
    }
}