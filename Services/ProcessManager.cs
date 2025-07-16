using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public interface IProcessManager
    {
        void KillOrphanedWordProcesses();
        void KillOrphanedExcelProcesses();
        bool IsProcessHealthy(int processId);
        Process[] GetWordProcesses();
        Process[] GetExcelProcesses();
        Task<bool> WaitForProcessExitAsync(int processId, TimeSpan timeout);
        void LogProcessInfo();
    }

    public class ProcessManager : IProcessManager, IDisposable
    {
        private readonly ILogger _logger;
        private readonly int _currentProcessId;
        private bool _disposed;

        public ProcessManager()
        {
            _logger = Log.ForContext<ProcessManager>();
            _currentProcessId = Process.GetCurrentProcess().Id;
            _logger.Information("ProcessManager initialized for parent process {ProcessId}", _currentProcessId);
        }

        public void KillOrphanedWordProcesses()
        {
            try
            {
                var wordProcesses = GetWordProcesses();
                var orphanedCount = 0;

                foreach (var process in wordProcesses)
                {
                    try
                    {
                        if (IsOrphanedProcess(process))
                        {
                            _logger.Warning("Killing orphaned Word process PID {ProcessId}", process.Id);
                            process.Kill();
                            process.WaitForExit(5000); // Wait up to 5 seconds
                            orphanedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to kill Word process PID {ProcessId}", process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }

                if (orphanedCount > 0)
                {
                    _logger.Information("Cleaned up {Count} orphaned Word processes", orphanedCount);
                }
                else
                {
                    _logger.Debug("No orphaned Word processes found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during Word process cleanup");
            }
        }

        public void KillOrphanedExcelProcesses()
        {
            try
            {
                var excelProcesses = GetExcelProcesses();
                var orphanedCount = 0;

                foreach (var process in excelProcesses)
                {
                    try
                    {
                        if (IsOrphanedProcess(process))
                        {
                            _logger.Warning("Killing orphaned Excel process PID {ProcessId}", process.Id);
                            process.Kill();
                            process.WaitForExit(5000); // Wait up to 5 seconds
                            orphanedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to kill Excel process PID {ProcessId}", process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }

                if (orphanedCount > 0)
                {
                    _logger.Information("Cleaned up {Count} orphaned Excel processes", orphanedCount);
                }
                else
                {
                    _logger.Debug("No orphaned Excel processes found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during Excel process cleanup");
            }
        }

        public bool IsProcessHealthy(int processId)
        {
            try
            {
                var process = Process.GetProcessById(processId);
                
                // Check if process is running and responsive
                bool isHealthy = !process.HasExited && 
                               process.Responding && 
                               !IsProcessHanging(process);
                
                process.Dispose();
                return isHealthy;
            }
            catch (ArgumentException)
            {
                // Process doesn't exist
                return false;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error checking process health for PID {ProcessId}", processId);
                return false;
            }
        }

        public Process[] GetWordProcesses()
        {
            try
            {
                return Process.GetProcessesByName("WINWORD");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error getting Word processes");
                return Array.Empty<Process>();
            }
        }

        public Process[] GetExcelProcesses()
        {
            try
            {
                return Process.GetProcessesByName("EXCEL");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error getting Excel processes");
                return Array.Empty<Process>();
            }
        }

        public async Task<bool> WaitForProcessExitAsync(int processId, TimeSpan timeout)
        {
            try
            {
                var process = Process.GetProcessById(processId);
                var exitedWithinTimeout = await Task.Run(() => process.WaitForExit((int)timeout.TotalMilliseconds));
                
                process.Dispose();
                return exitedWithinTimeout;
            }
            catch (ArgumentException)
            {
                // Process doesn't exist - consider it "exited"
                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error waiting for process exit PID {ProcessId}", processId);
                return false;
            }
        }

        public void LogProcessInfo()
        {
            try
            {
                var wordProcesses = GetWordProcesses();
                var excelProcesses = GetExcelProcesses();

                _logger.Information("Process Status - Word: {WordCount} processes, Excel: {ExcelCount} processes", 
                    wordProcesses.Length, excelProcesses.Length);

                // Log details for each Word process
                foreach (var process in wordProcesses)
                {
                    try
                    {
                        var isOrphaned = IsOrphanedProcess(process);
                        var memoryMB = process.WorkingSet64 / (1024 * 1024);
                        
                        _logger.Debug("Word Process PID {ProcessId}: Memory {MemoryMB}MB, Orphaned: {IsOrphaned}, Responding: {Responding}", 
                            process.Id, memoryMB, isOrphaned, process.Responding);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error logging Word process info for PID {ProcessId}", process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }

                // Log details for each Excel process
                foreach (var process in excelProcesses)
                {
                    try
                    {
                        var isOrphaned = IsOrphanedProcess(process);
                        var memoryMB = process.WorkingSet64 / (1024 * 1024);
                        
                        _logger.Debug("Excel Process PID {ProcessId}: Memory {MemoryMB}MB, Orphaned: {IsOrphaned}, Responding: {Responding}", 
                            process.Id, memoryMB, isOrphaned, process.Responding);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error logging Excel process info for PID {ProcessId}", process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error logging process information");
            }
        }

        private bool IsOrphanedProcess(Process process)
        {
            try
            {
                // Check if process has a parent or if parent is our current process
                var parentProcessId = GetParentProcessId(process);
                
                // Process is orphaned if:
                // 1. It has no parent (parent PID = 0)
                // 2. Its parent is our current process (we should manage cleanup)
                // 3. Its parent no longer exists
                if (parentProcessId == 0 || parentProcessId == _currentProcessId)
                {
                    return true;
                }

                // Check if parent process still exists
                try
                {
                    var parentProcess = Process.GetProcessById(parentProcessId);
                    var parentExists = !parentProcess.HasExited;
                    parentProcess.Dispose();
                    return !parentExists;
                }
                catch (ArgumentException)
                {
                    // Parent process doesn't exist
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error checking if process PID {ProcessId} is orphaned", process.Id);
                // When in doubt, assume it's not orphaned to avoid killing valid processes
                return false;
            }
        }

        private int GetParentProcessId(Process process)
        {
            try
            {
                // Use WMI to get parent process ID
                using var searcher = new System.Management.ManagementObjectSearcher(
                    $"SELECT ParentProcessId FROM Win32_Process WHERE ProcessId = {process.Id}");
                
                using var results = searcher.Get();
                foreach (System.Management.ManagementObject mo in results)
                {
                    return Convert.ToInt32(mo["ParentProcessId"]);
                }
                
                return 0; // No parent found
            }
            catch (Exception ex)
            {
                _logger.Debug(ex, "Error getting parent process ID for PID {ProcessId}", process.Id);
                return 0;
            }
        }

        private bool IsProcessHanging(Process process)
        {
            try
            {
                // Simple check: if process is not responding for more than 10 seconds
                // This is a basic implementation - could be enhanced with more sophisticated checks
                return !process.Responding;
            }
            catch (Exception)
            {
                return true; // Assume hanging if we can't check
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
                    // Log final process state
                    LogProcessInfo();
                    
                    // Optionally clean up orphaned processes on disposal
                    // Commented out for safety - only do this if explicitly requested
                    // KillOrphanedWordProcesses();
                    // KillOrphanedExcelProcesses();
                }
                
                _disposed = true;
                _logger.Information("ProcessManager disposed");
            }
        }

        ~ProcessManager()
        {
            Dispose(false);
        }
    }
} 