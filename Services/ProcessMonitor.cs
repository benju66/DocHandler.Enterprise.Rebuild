using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class ProcessMonitor : IDisposable
    {
        private readonly ILogger _logger;
        private readonly List<int> _trackedProcessIds = new();
        private readonly object _lock = new object();
        private bool _disposed = false;

        public ProcessMonitor(ILogger logger)
        {
            _logger = logger;
        }

        public void TrackProcess(int processId)
        {
            lock (_lock)
            {
                if (!_trackedProcessIds.Contains(processId))
                {
                    _trackedProcessIds.Add(processId);
                    _logger.Debug("Now tracking process {ProcessId}", processId);
                }
            }
        }

        public void UntrackProcess(int processId)
        {
            lock (_lock)
            {
                _trackedProcessIds.Remove(processId);
                _logger.Debug("Stopped tracking process {ProcessId}", processId);
            }
        }

        public async Task KillHungProcessesAsync(string processName)
        {
            try
            {
                await Task.Run(() =>
                {
                    var processes = Process.GetProcessesByName(processName);
                    foreach (var process in processes)
                    {
                        try
                        {
                            if (!process.HasExited && !process.Responding)
                            {
                                _logger.Warning("Killing hung {ProcessName} process with PID: {ProcessId}", 
                                    processName, process.Id);
                                process.Kill();
                                process.WaitForExit(5000);
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Failed to kill {ProcessName} process {ProcessId}", 
                                processName, process.Id);
                        }
                        finally
                        {
                            process.Dispose();
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error killing hung {ProcessName} processes", processName);
            }
        }

        public bool IsProcessRunning(int processId)
        {
            try
            {
                var process = Process.GetProcessById(processId);
                return !process.HasExited;
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                // Kill any tracked processes on disposal
                lock (_lock)
                {
                    foreach (var processId in _trackedProcessIds.ToList())
                    {
                        try
                        {
                            var process = Process.GetProcessById(processId);
                            if (!process.HasExited)
                            {
                                _logger.Warning("Killing tracked process {ProcessId} during disposal", processId);
                                process.Kill();
                            }
                        }
                        catch
                        {
                            // Process may already be gone
                        }
                    }
                    _trackedProcessIds.Clear();
                }
                
                _disposed = true;
                _logger.Information("ProcessMonitor disposed");
            }
        }
    }
} 