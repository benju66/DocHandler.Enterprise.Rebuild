using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Guards against orphaned Office processes by tracking and killing them on disposal
    /// </summary>
    public class OfficeProcessGuard : IDisposable
    {
        private readonly ILogger _logger;
        private readonly HashSet<int> _ourProcessIds = new HashSet<int>();
        private readonly HashSet<int> _preExistingProcessIds;
        private readonly object _lock = new object();
        private bool _disposed;

        public OfficeProcessGuard()
        {
            _logger = Log.ForContext<OfficeProcessGuard>();
            
            // Record all Office processes that existed before we started
            _preExistingProcessIds = new HashSet<int>();
            
            var wordProcesses = Process.GetProcessesByName("WINWORD");
            var excelProcesses = Process.GetProcessesByName("EXCEL");
            
            foreach (var process in wordProcesses.Concat(excelProcesses))
            {
                _preExistingProcessIds.Add(process.Id);
                process.Dispose();
            }
            
            _logger.Information("OfficeProcessGuard initialized. Pre-existing processes: {Count}", 
                _preExistingProcessIds.Count);
        }

        /// <summary>
        /// Register a process ID that we created
        /// </summary>
        public void RegisterProcess(int processId)
        {
            lock (_lock)
            {
                if (!_preExistingProcessIds.Contains(processId))
                {
                    _ourProcessIds.Add(processId);
                    _logger.Debug("Registered Office process: {PID}", processId);
                }
            }
        }

        /// <summary>
        /// Unregister a process ID (if it was closed normally)
        /// </summary>
        public void UnregisterProcess(int processId)
        {
            lock (_lock)
            {
                if (_ourProcessIds.Remove(processId))
                {
                    _logger.Debug("Unregistered Office process: {PID}", processId);
                }
            }
        }

        /// <summary>
        /// Kill all processes we created that are still running
        /// </summary>
        public void KillAllOurProcesses()
        {
            lock (_lock)
            {
                var killedCount = 0;
                
                foreach (var pid in _ourProcessIds.ToList())
                {
                    try
                    {
                        var process = Process.GetProcessById(pid);
                        if (!process.HasExited)
                        {
                            _logger.Warning("Force killing Office process {PID} ({Name})", 
                                pid, process.ProcessName);
                            
                            process.Kill();
                            process.WaitForExit(5000); // Wait up to 5 seconds
                            killedCount++;
                        }
                        process.Dispose();
                    }
                    catch (ArgumentException)
                    {
                        // Process no longer exists
                        _logger.Debug("Process {PID} no longer exists", pid);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to kill process {PID}", pid);
                    }
                }
                
                if (killedCount > 0)
                {
                    _logger.Warning("Killed {Count} orphaned Office processes", killedCount);
                }
                
                _ourProcessIds.Clear();
            }
        }

        /// <summary>
        /// Find any Office processes that we didn't create but might be orphaned
        /// </summary>
        public List<Process> FindPotentiallyOrphanedProcesses()
        {
            var orphaned = new List<Process>();
            
            try
            {
                var wordProcesses = Process.GetProcessesByName("WINWORD");
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                
                foreach (var process in wordProcesses.Concat(excelProcesses))
                {
                    // If it's not pre-existing and not one we registered, it might be orphaned
                    if (!_preExistingProcessIds.Contains(process.Id) && 
                        !_ourProcessIds.Contains(process.Id))
                    {
                        // Check if it has a main window (user-created processes usually do)
                        if (process.MainWindowHandle == IntPtr.Zero)
                        {
                            orphaned.Add(process);
                        }
                        else
                        {
                            process.Dispose();
                        }
                    }
                    else
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error finding orphaned processes");
            }
            
            return orphaned;
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
                    // Kill all processes we created
                    KillAllOurProcesses();
                }
                
                _disposed = true;
                _logger.Information("OfficeProcessGuard disposed");
            }
        }
    }
} 