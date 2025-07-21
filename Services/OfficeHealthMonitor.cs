using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Monitors Office health and automatically recovers from memory leaks
    /// </summary>
    public class OfficeHealthMonitor : IDisposable
    {
        private readonly ILogger _logger;
        private readonly Timer _monitorTimer;
        private readonly Action<HealthCheckResult> _recoveryAction;
        private bool _disposed;
        
        // Thresholds
        private readonly int _maxComObjects = 10;
        private readonly long _maxMemoryBytes = 1_000_000_000; // 1GB
        private readonly TimeSpan _checkInterval = TimeSpan.FromMinutes(1);
        
        public class HealthCheckResult
        {
            public long ComObjectCount { get; set; }
            public long MemoryBytes { get; set; }
            public bool IsHealthy { get; set; }
            public string Reason { get; set; }
        }
        
        public OfficeHealthMonitor(Action<HealthCheckResult> recoveryAction = null)
        {
            _logger = Log.ForContext<OfficeHealthMonitor>();
            _recoveryAction = recoveryAction;
            
            _monitorTimer = new Timer(PerformHealthCheck, null, _checkInterval, _checkInterval);
            _logger.Information("Office health monitor started - checking every {Minutes} minutes", 
                _checkInterval.TotalMinutes);
        }
        
        private void PerformHealthCheck(object state)
        {
            try
            {
                var result = CheckHealth();
                
                if (!result.IsHealthy)
                {
                    _logger.Warning("Health check failed: {Reason}", result.Reason);
                    _logger.Warning("COM Objects: {ComObjects}, Memory: {MemoryMB}MB", 
                        result.ComObjectCount, result.MemoryBytes / 1_000_000);
                    
                    // Invoke recovery action if provided
                    _recoveryAction?.Invoke(result);
                    
                    // Force cleanup
                    ForceCleanup();
                }
                else
                {
                    _logger.Debug("Health check passed - COM: {ComObjects}, Memory: {MemoryMB}MB", 
                        result.ComObjectCount, result.MemoryBytes / 1_000_000);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during health check");
            }
        }
        
        public HealthCheckResult CheckHealth()
        {
            var result = new HealthCheckResult();
            
            try
            {
                // Check COM objects
                var stats = ComHelper.GetComObjectStats();
                result.ComObjectCount = stats.NetObjects;
                
                // Check memory
                using (var process = Process.GetCurrentProcess())
                {
                    result.MemoryBytes = process.WorkingSet64;
                }
                
                // Determine health
                if (result.ComObjectCount > _maxComObjects)
                {
                    result.IsHealthy = false;
                    result.Reason = $"High COM object count: {result.ComObjectCount} (max: {_maxComObjects})";
                }
                else if (result.MemoryBytes > _maxMemoryBytes)
                {
                    result.IsHealthy = false;
                    result.Reason = $"High memory usage: {result.MemoryBytes / 1_000_000}MB (max: {_maxMemoryBytes / 1_000_000}MB)";
                }
                else
                {
                    result.IsHealthy = true;
                    result.Reason = "All metrics within normal range";
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error checking health metrics");
                result.IsHealthy = false;
                result.Reason = "Error checking health: " + ex.Message;
            }
            
            return result;
        }
        
        private void ForceCleanup()
        {
            try
            {
                _logger.Information("Forcing COM cleanup due to health check failure");
                
                // Force garbage collection
                ComHelper.ForceComCleanup("HealthMonitor");
                
                // Log updated stats
                ComHelper.LogComObjectStats();
                
                // Check for orphaned processes
                CheckForOrphanedProcesses();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during forced cleanup");
            }
        }
        
        private void CheckForOrphanedProcesses()
        {
            try
            {
                var wordProcesses = Process.GetProcessesByName("WINWORD");
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                
                var orphanedCount = 0;
                
                foreach (var process in wordProcesses.Concat(excelProcesses))
                {
                    try
                    {
                        // Orphaned processes typically have no main window
                        if (process.MainWindowHandle == IntPtr.Zero)
                        {
                            _logger.Warning("Found potential orphaned {ProcessName} process (PID: {PID})", 
                                process.ProcessName, process.Id);
                            orphanedCount++;
                        }
                    }
                    catch { }
                    finally
                    {
                        process.Dispose();
                    }
                }
                
                if (orphanedCount > 0)
                {
                    _logger.Warning("Found {Count} potential orphaned Office processes", orphanedCount);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error checking for orphaned processes");
            }
        }
        
        public void Stop()
        {
            _monitorTimer?.Change(Timeout.Infinite, Timeout.Infinite);
            _logger.Information("Health monitor stopped");
        }
        
        public void Dispose()
        {
            if (!_disposed)
            {
                Stop();
                _monitorTimer?.Dispose();
                _disposed = true;
                _logger.Information("Health monitor disposed");
            }
        }
    }
} 