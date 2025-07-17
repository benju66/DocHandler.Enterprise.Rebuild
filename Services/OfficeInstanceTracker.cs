using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Tracks Office instances to distinguish between user-created and app-created instances
    /// Prevents closing user's existing Office applications on shutdown
    /// </summary>
    public class OfficeInstanceTracker : IDisposable
    {
        private readonly ILogger _logger;
        private readonly ConcurrentHashSet<int> _preExistingWordProcesses = new();
        private readonly ConcurrentHashSet<int> _preExistingExcelProcesses = new();
        private readonly ConcurrentHashSet<int> _appCreatedWordProcesses = new();
        private readonly ConcurrentHashSet<int> _appCreatedExcelProcesses = new();
        private readonly object _lockObject = new object();
        private bool _disposed = false;

        public OfficeInstanceTracker()
        {
            _logger = Log.ForContext<OfficeInstanceTracker>();
            RecordPreExistingInstances();
        }

        /// <summary>
        /// Records all currently running Office instances as pre-existing (user-owned)
        /// </summary>
        private void RecordPreExistingInstances()
        {
            lock (_lockObject)
            {
                try
                {
                    // Record existing Word processes
                    var wordProcesses = Process.GetProcessesByName("WINWORD");
                    foreach (var process in wordProcesses)
                    {
                        _preExistingWordProcesses.Add(process.Id);
                        _logger.Information("Recorded pre-existing Word process PID {ProcessId} (USER-OWNED)", process.Id);
                        process.Dispose();
                    }

                    // Record existing Excel processes  
                    var excelProcesses = Process.GetProcessesByName("EXCEL");
                    foreach (var process in excelProcesses)
                    {
                        _preExistingExcelProcesses.Add(process.Id);
                        _logger.Information("Recorded pre-existing Excel process PID {ProcessId} (USER-OWNED)", process.Id);
                        process.Dispose();
                    }

                    _logger.Information("Instance tracking initialized: {WordCount} Word, {ExcelCount} Excel processes marked as user-owned",
                        _preExistingWordProcesses.Count, _preExistingExcelProcesses.Count);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to record pre-existing Office instances");
                }
            }
        }

        /// <summary>
        /// Registers a new Word process as app-created
        /// </summary>
        public void RegisterAppCreatedWordProcess(int processId)
        {
            lock (_lockObject)
            {
                if (!_preExistingWordProcesses.Contains(processId))
                {
                    _appCreatedWordProcesses.Add(processId);
                    _logger.Information("Registered app-created Word process PID {ProcessId} (APP-OWNED)", processId);
                }
                else
                {
                    _logger.Warning("Process PID {ProcessId} was pre-existing, not registering as app-created", processId);
                }
            }
        }

        /// <summary>
        /// Registers a new Excel process as app-created
        /// </summary>
        public void RegisterAppCreatedExcelProcess(int processId)
        {
            lock (_lockObject)
            {
                if (!_preExistingExcelProcesses.Contains(processId))
                {
                    _appCreatedExcelProcesses.Add(processId);
                    _logger.Information("Registered app-created Excel process PID {ProcessId} (APP-OWNED)", processId);
                }
                else
                {
                    _logger.Warning("Process PID {ProcessId} was pre-existing, not registering as app-created", processId);
                }
            }
        }

        /// <summary>
        /// Unregisters an app-created process (when properly disposed)
        /// </summary>
        public void UnregisterAppCreatedWordProcess(int processId)
        {
            lock (_lockObject)
            {
                if (_appCreatedWordProcesses.Remove(processId))
                {
                    _logger.Information("Unregistered app-created Word process PID {ProcessId}", processId);
                }
            }
        }

        /// <summary>
        /// Unregisters an app-created process (when properly disposed)
        /// </summary>
        public void UnregisterAppCreatedExcelProcess(int processId)
        {
            lock (_lockObject)
            {
                if (_appCreatedExcelProcesses.Remove(processId))
                {
                    _logger.Information("Unregistered app-created Excel process PID {ProcessId}", processId);
                }
            }
        }

        /// <summary>
        /// Checks if a Word process was created by the application
        /// </summary>
        public bool IsAppCreatedWordProcess(int processId)
        {
            lock (_lockObject)
            {
                return _appCreatedWordProcesses.Contains(processId);
            }
        }

        /// <summary>
        /// Checks if an Excel process was created by the application
        /// </summary>
        public bool IsAppCreatedExcelProcess(int processId)
        {
            lock (_lockObject)
            {
                return _appCreatedExcelProcesses.Contains(processId);
            }
        }

        /// <summary>
        /// Checks if a Word process is user-owned (pre-existing)
        /// </summary>
        public bool IsUserOwnedWordProcess(int processId)
        {
            lock (_lockObject)
            {
                return _preExistingWordProcesses.Contains(processId);
            }
        }

        /// <summary>
        /// Checks if an Excel process is user-owned (pre-existing)  
        /// </summary>
        public bool IsUserOwnedExcelProcess(int processId)
        {
            lock (_lockObject)
            {
                return _preExistingExcelProcesses.Contains(processId);
            }
        }

        /// <summary>
        /// Gets all app-created Word processes that should be cleaned up
        /// </summary>
        public List<int> GetAppCreatedWordProcesses()
        {
            lock (_lockObject)
            {
                return _appCreatedWordProcesses.ToList();
            }
        }

        /// <summary>
        /// Gets all app-created Excel processes that should be cleaned up
        /// </summary>
        public List<int> GetAppCreatedExcelProcesses()
        {
            lock (_lockObject)
            {
                return _appCreatedExcelProcesses.ToList();
            }
        }

        /// <summary>
        /// Safely cleans up only app-created Office processes
        /// Enhanced safety: never kill processes unless absolutely certain they belong to the app
        /// </summary>
        public void CleanupAppCreatedProcesses()
        {
            lock (_lockObject)
            {
                _logger.Information("Starting safe cleanup of app-created Office processes");

                // Cleanup app-created Word processes with extra safety checks
                var wordProcessesToCleanup = _appCreatedWordProcesses.ToList();
                foreach (var processId in wordProcessesToCleanup)
                {
                    try
                    {
                        // Multiple safety checks before killing any process
                        if (!IsDefinitelyAppCreated(processId, _appCreatedWordProcesses, _preExistingWordProcesses))
                        {
                            _logger.Warning("Skipping process PID {ProcessId} - safety check failed", processId);
                            continue;
                        }

                        var process = Process.GetProcessById(processId);
                        if (!process.HasExited)
                        {
                            _logger.Information("Gracefully closing app-created Word process PID {ProcessId}", processId);
                            
                            // Try graceful shutdown first
                            if (!process.CloseMainWindow())
                            {
                                _logger.Debug("CloseMainWindow failed for process {ProcessId}, trying WM_CLOSE", processId);
                                // Send WM_CLOSE message as backup
                                try
                                {
                                    SendCloseMessage(process.MainWindowHandle);
                                }
                                catch (Exception ex)
                                {
                                    _logger.Debug(ex, "Could not send close message to process {ProcessId}", processId);
                                }
                            }
                            
                            // Wait for graceful exit
                            if (!process.WaitForExit(10000)) // 10 second timeout
                            {
                                _logger.Warning("Process {ProcessId} did not close gracefully, forcing termination", processId);
                                process.Kill();
                                process.WaitForExit(5000);
                            }
                            else
                            {
                                _logger.Information("Process {ProcessId} closed gracefully", processId);
                            }
                        }
                        _appCreatedWordProcesses.Remove(processId);
                        process.Dispose();
                    }
                    catch (ArgumentException)
                    {
                        // Process already gone - remove from tracking
                        _appCreatedWordProcesses.Remove(processId);
                        _logger.Debug("Process PID {ProcessId} already exited", processId);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error during safe cleanup of Word process PID {ProcessId}", processId);
                        // Don't remove from tracking if we couldn't clean it up
                    }
                }

                // Cleanup app-created Excel processes with same safety measures
                var excelProcessesToCleanup = _appCreatedExcelProcesses.ToList();
                foreach (var processId in excelProcessesToCleanup)
                {
                    try
                    {
                        // Multiple safety checks before killing any process
                        if (!IsDefinitelyAppCreated(processId, _appCreatedExcelProcesses, _preExistingExcelProcesses))
                        {
                            _logger.Warning("Skipping Excel process PID {ProcessId} - safety check failed", processId);
                            continue;
                        }

                        var process = Process.GetProcessById(processId);
                        if (!process.HasExited)
                        {
                            _logger.Information("Gracefully closing app-created Excel process PID {ProcessId}", processId);
                            
                            // Try graceful shutdown first
                            if (!process.CloseMainWindow())
                            {
                                _logger.Debug("CloseMainWindow failed for Excel process {ProcessId}, trying WM_CLOSE", processId);
                                try
                                {
                                    SendCloseMessage(process.MainWindowHandle);
                                }
                                catch (Exception ex)
                                {
                                    _logger.Debug(ex, "Could not send close message to Excel process {ProcessId}", processId);
                                }
                            }
                            
                            // Wait for graceful exit
                            if (!process.WaitForExit(10000)) // 10 second timeout
                            {
                                _logger.Warning("Excel process {ProcessId} did not close gracefully, forcing termination", processId);
                                process.Kill();
                                process.WaitForExit(5000);
                            }
                            else
                            {
                                _logger.Information("Excel process {ProcessId} closed gracefully", processId);
                            }
                        }
                        _appCreatedExcelProcesses.Remove(processId);
                        process.Dispose();
                    }
                    catch (ArgumentException)
                    {
                        // Process already gone - remove from tracking
                        _appCreatedExcelProcesses.Remove(processId);
                        _logger.Debug("Excel process PID {ProcessId} already exited", processId);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error during safe cleanup of Excel process PID {ProcessId}", processId);
                        // Don't remove from tracking if we couldn't clean it up
                    }
                }

                _logger.Information("Safe cleanup completed. Processed {WordCount} Word and {ExcelCount} Excel app-created processes",
                    wordProcessesToCleanup.Count, excelProcessesToCleanup.Count);

                // Log remaining user processes (should not be touched)
                var remainingWord = Process.GetProcessesByName("WINWORD");
                var remainingExcel = Process.GetProcessesByName("EXCEL");
                _logger.Information("User Office processes preserved: {WordCount} Word, {ExcelCount} Excel",
                    remainingWord.Length, remainingExcel.Length);

                foreach (var process in remainingWord) process.Dispose();
                foreach (var process in remainingExcel) process.Dispose();
            }
        }

        /// <summary>
        /// Performs multiple safety checks to ensure a process is definitely app-created
        /// </summary>
        private bool IsDefinitelyAppCreated(int processId, ConcurrentHashSet<int> appCreated, ConcurrentHashSet<int> preExisting)
        {
            // Safety check 1: Must be in app-created list
            if (!appCreated.Contains(processId))
            {
                _logger.Debug("Process {ProcessId} not in app-created list", processId);
                return false;
            }

            // Safety check 2: Must NOT be in pre-existing list
            if (preExisting.Contains(processId))
            {
                _logger.Warning("Process {ProcessId} was pre-existing - should not kill!", processId);
                return false;
            }

            // Safety check 3: Process must exist and be accessible
            try
            {
                var process = Process.GetProcessById(processId);
                var canAccess = !process.HasExited;
                process.Dispose();
                
                if (!canAccess)
                {
                    _logger.Debug("Process {ProcessId} has already exited", processId);
                    return false;
                }
            }
            catch (ArgumentException)
            {
                _logger.Debug("Process {ProcessId} does not exist", processId);
                return false;
            }
            catch (Exception ex)
            {
                _logger.Debug(ex, "Cannot access process {ProcessId}", processId);
                return false;
            }

            return true;
        }

        /// <summary>
        /// Sends WM_CLOSE message to a window handle for graceful shutdown
        /// </summary>
        private void SendCloseMessage(IntPtr windowHandle)
        {
            if (windowHandle != IntPtr.Zero)
            {
                const int WM_CLOSE = 0x0010;
                SendMessage(windowHandle, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        public void Dispose()
        {
            if (!_disposed)
            {
                CleanupAppCreatedProcesses();
                _disposed = true;
                _logger.Information("OfficeInstanceTracker disposed");
            }
        }
    }

    /// <summary>
    /// Thread-safe HashSet implementation
    /// </summary>
    public class ConcurrentHashSet<T> : IDisposable
    {
        private readonly HashSet<T> _hashSet = new HashSet<T>();
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim();

        public void Add(T item)
        {
            _lock.EnterWriteLock();
            try
            {
                _hashSet.Add(item);
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        public bool Remove(T item)
        {
            _lock.EnterWriteLock();
            try
            {
                return _hashSet.Remove(item);
            }
            finally
            {
                _lock.ExitWriteLock();
            }
        }

        public bool Contains(T item)
        {
            _lock.EnterReadLock();
            try
            {
                return _hashSet.Contains(item);
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        public List<T> ToList()
        {
            _lock.EnterReadLock();
            try
            {
                return _hashSet.ToList();
            }
            finally
            {
                _lock.ExitReadLock();
            }
        }

        public int Count
        {
            get
            {
                _lock.EnterReadLock();
                try
                {
                    return _hashSet.Count;
                }
                finally
                {
                    _lock.ExitReadLock();
                }
            }
        }

        public void Dispose()
        {
            _lock?.Dispose();
        }
    }
} 