using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Serilog;
using System.Runtime.InteropServices;
using System.Threading;

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
        private bool _isWarmingUp = false;
        private readonly object _warmUpLock = new object();
        
        // Add idle timer mechanism
        private DateTime _lastUsed = DateTime.Now;
        private Timer? _idleTimer;
        private readonly TimeSpan _idleTimeout = TimeSpan.FromMinutes(5);
        
        // Diagnostic tracking fields
        private DateTime _createdAt;
        private bool _wasCreatedByUs = false;
        
        public SessionAwareExcelService()
        {
            _logger = Log.ForContext<SessionAwareExcelService>();
        }
        
        private void InitializeExcel()
        {
            try
            {
                // Validate STA thread before COM operations
                var apartmentState = Thread.CurrentThread.GetApartmentState();
                if (apartmentState != ApartmentState.STA)
                {
                    _logger.Error("InitializeExcel: Thread is not STA ({ApartmentState}) - COM operations may fail", apartmentState);
                    throw new InvalidOperationException($"Thread must be STA for COM operations. Current state: {apartmentState}");
                }
                
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    throw new InvalidOperationException("Excel.Application ProgID not found");
                }
                
                _excelApp = Activator.CreateInstance(excelType);
                ComHelper.TrackComObjectCreation("ExcelApp", "SessionAwareExcelService");
                
                // Set diagnostic tracking
                _createdAt = DateTime.Now;
                _wasCreatedByUs = true;
                
                _logger.Information("Excel app created - Thread: {ThreadId}, Time: {Time}", 
                    System.Threading.Thread.CurrentThread.ManagedThreadId, DateTime.Now);
                    
                _excelApp.Visible = false;
                _excelApp.DisplayAlerts = false;
                _excelApp.ScreenUpdating = false;
                
                // Do not set WindowState - it can cause Excel to briefly flash visible
                
                _lastHealthCheck = DateTime.Now;
                _lastUsed = DateTime.Now;
                
                // Start idle timer
                _idleTimer = new Timer(CheckIdleTimeout, null, _idleTimeout, _idleTimeout);
                
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
                // Don't access any properties - just check if not null
                // The real test happens when we try to use it
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
            // CRITICAL FIX: Remove Task.Run - already on STA thread from caller
            lock (_conversionLock)
            {
                    _lastUsed = DateTime.Now;
                    
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
                        using (var comScope = new ComResourceScope())
                        {
                            workbook = comScope.OpenExcelWorkbook(_excelApp, inputPath, readOnly: true);
                            // ComResourceScope already tracks the workbook - removed duplicate tracking
                            workbook.ExportAsFixedFormat(
                                Type: 0, // xlTypePDF
                                Filename: outputPath,
                                Quality: 0); // xlQualityStandard
                        }
                        
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
                                // CRITICAL FIX: Don't release workbook here - ComResourceScope already handled it
                                // ComHelper.SafeReleaseComObject(workbook, "Workbook", "SessionAwareConvertToPdf");
                            }
                            catch (Exception ex)
                            {
                                _logger.Warning(ex, "Error closing Excel workbook");
                            }
                        }
                    }
                    
                    return result;
                }
        }
        
        public void WarmUp()
        {
            lock (_warmUpLock)
            {
                if (_isWarmingUp) return; // Prevent multiple warm-ups
                _isWarmingUp = true;
            }
            
            try
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
            finally
            {
                lock (_warmUpLock)
                {
                    _isWarmingUp = false;
                }
            }
        }
        
        // Add method to dispose when Save Quotes Mode is disabled
        public void DisposeIfIdle()
        {
            lock (_conversionLock)
            {
                if (_excelApp != null)
                {
                    _logger.Information("Disposing idle Excel instance");
                    DisposeExcel();
                }
            }
        }
        
        private void DisposeExcel()
        {
            _logger.Information("DisposeExcel called, _excelApp is {Status}", _excelApp != null ? "not null" : "null");
            
            // CRITICAL FIX: Dispose timer first to prevent callbacks during disposal
            if (_idleTimer != null)
            {
                try
                {
                    _idleTimer.Dispose();
                    _idleTimer = null;
                    _logger.Debug("Idle timer disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing idle timer");
                }
            }
            
            try
            {
                if (_excelApp != null)
                {
                    try
                    {
                        // Diagnostic logging before disposal
                        int processId = 0;
                        bool isVisible = false;
                        try 
                        {
                            // Try to get process ID (Excel doesn't have the safe method like Word)
                            var process = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                                .FirstOrDefault(p => !p.HasExited);
                            if (process != null)
                            {
                                processId = process.Id;
                                process.Dispose();
                            }
                            isVisible = _excelApp.Visible;
                        }
                        catch { }
                        
                        _logger.Information("Disposing Excel - PID: {ProcessId}, Visible: {Visible}, CreatedByUs: {CreatedByUs}, CreatedAt: {CreatedAt}", 
                            processId, isVisible, _wasCreatedByUs, _createdAt);
                        
                        // Try to close all workbooks first
                        try
                        {
                            using (var comScope = new ComResourceScope())
                            {
                                var workbooks = comScope.GetWorkbooks(_excelApp, "DisposeExcel");
                                if (workbooks != null && workbooks.Count > 0)
                                {
                                    _logger.Warning("Closing {Count} open workbooks", workbooks.Count);
                                    // Use for loop instead of foreach to avoid COM enumerator leak
                                    int count = workbooks.Count;
                                    for (int i = count; i >= 1; i--) // Excel collections are 1-based, iterate backwards
                                    {
                                        try
                                        {
                                            dynamic wb = workbooks[i];
                                            wb.Close(false);
                                            comScope.Track(wb, "Workbook", "DisposeExcel");
                                        }
                                        catch { }
                                    }
                                }
                            } // workbooks collection automatically released here
                        }
                        catch { }
                        
                        // Enhanced Quit() with better error handling
                        try
                        {
                            _excelApp.Quit();
                            _logger.Information("Excel Quit() completed successfully");
                        }
                        catch (COMException comEx)
                        {
                            _logger.Warning(comEx, "COM exception during Excel Quit() - app may be disconnected (HRESULT: 0x{HResult:X8})", comEx.HResult);
                        }
                        catch (Exception quitEx)
                        {
                            _logger.Warning(quitEx, "Unexpected exception during Excel Quit()");
                        }
                        
                        ComHelper.SafeReleaseComObject(_excelApp, "ExcelApp", "DisposeExcel");
                        _logger.Information("Excel COM object released and set to null");
                        _excelApp = null;
                        
                        // Reset tracking
                        _wasCreatedByUs = false;
                        _createdAt = default;
                        
                        // Force garbage collection
                        ComHelper.ForceComCleanup("SessionAwareExcelService");
                        
                        _logger.Information("Excel disposal completed successfully - Thread: {ThreadId}, Time: {Time}", 
                            System.Threading.Thread.CurrentThread.ManagedThreadId, DateTime.Now);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error disposing Excel application");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error in DisposeExcel");
            }
        }
        
        private void CheckIdleTimeout(object? state)
        {
            lock (_conversionLock)
            {
                if (_excelApp != null && DateTime.Now - _lastUsed > _idleTimeout)
                {
                    _logger.Information("Excel application idle for {Minutes} minutes, disposing", _idleTimeout.TotalMinutes);
                    DisposeExcel();
                    
                    // CRITICAL FIX: Ensure timer is cleared after disposing Excel to prevent further callbacks
                    _idleTimer = null;
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
            _logger.Information("SessionAwareExcelService.Dispose called with disposing={Disposing}", disposing);
            
            if (!_disposed)
            {
                if (disposing)
                {
                    // Dispose managed resources
                    _idleTimer?.Dispose();
                }
                
                // Dispose unmanaged resources (COM objects)
                DisposeExcel();
                
                _disposed = true;
            }
        }
        
        // CRITICAL FIX: Finalizer ensures COM objects are released even if Dispose isn't called
        ~SessionAwareExcelService()
        {
            Dispose(false);
        }
    }
} 