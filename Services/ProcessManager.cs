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
                
                // THREADING FIX: Use process.WaitForExitAsync if available (.NET 5+)
                // For older frameworks, use a more efficient polling approach instead of Task.Run
                try
                {
                    using var cts = new System.Threading.CancellationTokenSource(timeout);
                    await process.WaitForExitAsync(cts.Token);
                    process.Dispose();
                    return true;
                }
                catch (OperationCanceledException)
                {
                    // Timeout occurred
                    process.Dispose();
                    return false;
                }
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
                
                // Add these variable declarations
                var wordCount = wordProcesses.Length;
                var excelCount = excelProcesses.Length;

                _logger.Information("Process Status - Word: {WordCount} processes, Excel: {ExcelCount} processes", 
                    wordCount, excelCount);

                // Only log process details if there are issues or in verbose mode
                var orphanedWordCount = 0;
                var highMemoryWordCount = 0;
                
                foreach (var process in wordProcesses)
                {
                    try
                    {
                        var isOrphaned = IsOrphanedProcess(process);
                        var memoryMB = process.WorkingSet64 / (1024 * 1024);
                        
                        if (isOrphaned) orphanedWordCount++;
                        if (memoryMB > 500) highMemoryWordCount++; // Flag processes using > 500MB
                        
                        // Only log individual processes if they're problematic
                        if (isOrphaned || memoryMB > 500 || !process.Responding)
                        {
                            _logger.Warning("Problematic Word Process PID {ProcessId}: Memory {MemoryMB}MB, Orphaned: {IsOrphaned}, Responding: {Responding}", 
                                process.Id, memoryMB, isOrphaned, process.Responding);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Debug(ex, "Error checking Word process info for PID {ProcessId}", process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
                
                // Log summary instead of individual processes
                if (wordCount > 0)
                {
                    _logger.Debug("Word processes: {Total} total, {Orphaned} orphaned, {HighMemory} high memory", 
                        wordCount, orphanedWordCount, highMemoryWordCount);
                }

                // Only log Excel process details if there are issues
                var orphanedExcelCount = 0;
                var highMemoryExcelCount = 0;
                
                foreach (var process in excelProcesses)
                {
                    try
                    {
                        var isOrphaned = IsOrphanedProcess(process);
                        var memoryMB = process.WorkingSet64 / (1024 * 1024);
                        
                        if (isOrphaned) orphanedExcelCount++;
                        if (memoryMB > 500) highMemoryExcelCount++;
                        
                        // Only log individual processes if they're problematic
                        if (isOrphaned || memoryMB > 500 || !process.Responding)
                        {
                            _logger.Warning("Problematic Excel Process PID {ProcessId}: Memory {MemoryMB}MB, Orphaned: {IsOrphaned}, Responding: {Responding}", 
                                process.Id, memoryMB, isOrphaned, process.Responding);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Debug(ex, "Error checking Excel process info for PID {ProcessId}", process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
                
                // Log summary for Excel processes
                if (excelCount > 0)
                {
                    _logger.Debug("Excel processes: {Total} total, {Orphaned} orphaned, {HighMemory} high memory", 
                        excelCount, orphanedExcelCount, highMemoryExcelCount);
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

        /// <summary>
        /// Terminate any orphaned Office processes that might have been left behind
        /// </summary>
        public void TerminateOrphanedOfficeProcesses()
        {
            try
            {
                _logger.Information("Checking for orphaned Office processes...");
                
                var killedCount = 0;
                var wordProcesses = Process.GetProcessesByName("WINWORD");
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                
                foreach (var process in wordProcesses.Concat(excelProcesses))
                {
                    try
                    {
                        // Check if it's likely an automated/orphaned process:
                        // - No main window (automated instances don't have UI)
                        // - Was started after our app (check process start time)
                        if (process.MainWindowHandle == IntPtr.Zero)
                        {
                            try
                            {
                                var currentProcess = Process.GetCurrentProcess();
                                if (process.StartTime > currentProcess.StartTime)
                                {
                                    _logger.Warning("Killing likely orphaned {ProcessName} process (PID: {PID}, StartTime: {StartTime})", 
                                        process.ProcessName, process.Id, process.StartTime);
                                    
                                    process.Kill();
                                    process.WaitForExit(5000); // Wait up to 5 seconds
                                    killedCount++;
                                }
                                currentProcess.Dispose();
                            }
                            catch (Exception ex)
                            {
                                _logger.Debug(ex, "Could not check process start time for PID {PID}", process.Id);
                            }
                        }
                        process.Dispose();
                    }
                    catch (Exception ex)
                    {
                        _logger.Debug(ex, "Error checking process {PID}", process.Id);
                    }
                }
                
                if (killedCount > 0)
                {
                    _logger.Warning("Terminated {Count} orphaned Office processes", killedCount);
                }
                else
                {
                    _logger.Information("No orphaned Office processes found");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error terminating orphaned Office processes");
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