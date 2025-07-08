using System;
using System.IO;
using System.Threading.Tasks;
using Serilog;
using System.Runtime.InteropServices;

namespace DocHandler.Services
{
    public class SessionAwareExcelService : IDisposable
    {
        private readonly ILogger _logger;
        private readonly object _conversionLock = new object();
        private dynamic? _excelApp;
        private bool _disposed;
        private DateTime _lastHealthCheck = DateTime.Now;
        private readonly TimeSpan _healthCheckInterval = TimeSpan.FromMinutes(5);
        
        public SessionAwareExcelService()
        {
            _logger = Log.ForContext<SessionAwareExcelService>();
        }
        
        private void InitializeExcel()
        {
            try
            {
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel.Application ProgID not found");
                }
                
                _excelApp = Activator.CreateInstance(excelType);
                _excelApp.Visible = false;
                _excelApp.DisplayAlerts = false;
                _excelApp.ScreenUpdating = false;
                
                // Minimize window
                _excelApp.WindowState = -4137; // xlMinimized
                
                _lastHealthCheck = DateTime.Now;
                _logger.Information("Excel application initialized for session (hidden mode)");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to initialize Excel application");
                throw;
            }
        }
        
        private bool IsExcelHealthy()
        {
            try
            {
                if (_excelApp == null) return false;
                var _ = _excelApp.Version;
                return true;
            }
            catch
            {
                _logger.Warning("Excel health check failed");
                return false;
            }
        }
        
        private void EnsureExcelHealthy()
        {
            if (DateTime.Now - _lastHealthCheck > _healthCheckInterval)
            {
                if (!IsExcelHealthy())
                {
                    _logger.Warning("Excel unhealthy - reinitializing");
                    DisposeExcel();
                    InitializeExcel();
                }
                _lastHealthCheck = DateTime.Now;
            }
        }
        
        public async Task<ConversionResult> ConvertSpreadsheetToPdf(string inputPath, string outputPath)
        {
            return await Task.Run(() =>
            {
                lock (_conversionLock)
                {
                    if (_excelApp == null)
                    {
                        InitializeExcel();
                    }
                    else
                    {
                        EnsureExcelHealthy();
                    }
                    
                    var result = new ConversionResult { Success = true };
                    dynamic? workbook = null;
                    
                    try
                    {
                        workbook = _excelApp.Workbooks.Open(inputPath, ReadOnly: true);
                        workbook.ExportAsFixedFormat(
                            Type: 0, // xlTypePDF
                            Filename: outputPath,
                            Quality: 0); // xlQualityStandard
                        
                        result.OutputPath = outputPath;
                        _logger.Debug("Converted Excel to PDF: {Input} -> {Output}", 
                            Path.GetFileName(inputPath), Path.GetFileName(outputPath));
                    }
                    catch (Exception ex)
                    {
                        result.Success = false;
                        result.ErrorMessage = $"Excel conversion failed: {ex.Message}";
                        _logger.Error(ex, "Failed to convert Excel file: {Path}", inputPath);
                    }
                    finally
                    {
                        if (workbook != null)
                        {
                            try
                            {
                                workbook.Close(false);
                                Marshal.ReleaseComObject(workbook);
                            }
                            catch (Exception ex)
                            {
                                _logger.Warning(ex, "Error closing Excel workbook");
                            }
                        }
                    }
                    
                    return result;
                }
            });
        }
        
        public void WarmUp()
        {
            lock (_conversionLock)
            {
                if (_excelApp == null)
                {
                    InitializeExcel();
                    _logger.Information("Excel pre-warmed for Save Quotes Mode");
                }
            }
        }
        
        private void DisposeExcel()
        {
            try
            {
                if (_excelApp != null)
                {
                    try
                    {
                        // Try to close all workbooks first
                        var workbooks = _excelApp.Workbooks;
                        if (workbooks != null && workbooks.Count > 0)
                        {
                            _logger.Warning("Closing {Count} open workbooks", workbooks.Count);
                            foreach (dynamic wb in workbooks)
                            {
                                try
                                {
                                    wb.Close(false);
                                    Marshal.ReleaseComObject(wb);
                                }
                                catch { }
                            }
                            Marshal.ReleaseComObject(workbooks);
                        }
                    }
                    catch { }
                    
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                    _excelApp = null;
                    
                    // Force garbage collection
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    
                    _logger.Information("Excel application disposed");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error disposing Excel application");
            }
        }
        
        public void Dispose()
        {
            if (!_disposed)
            {
                DisposeExcel();
                _disposed = true;
                GC.SuppressFinalize(this);
            }
        }
    }
} 