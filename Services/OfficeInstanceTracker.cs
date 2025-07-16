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
        /// </summary>
        public void CleanupAppCreatedProcesses()
        {
            lock (_lockObject)
            {
                _logger.Information("Starting cleanup of app-created Office processes");

                // Cleanup app-created Word processes
                var wordProcessesToCleanup = _appCreatedWordProcesses.ToList();
                foreach (var processId in wordProcessesToCleanup)
                {
                    try
                    {
                        var process = Process.GetProcessById(processId);
                        if (!process.HasExited)
                        {
                            _logger.Information("Terminating app-created Word process PID {ProcessId}", processId);
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                        _appCreatedWordProcesses.Remove(processId);
                        process.Dispose();
                    }
                    catch (ArgumentException)
                    {
                        // Process already gone
                        _appCreatedWordProcesses.Remove(processId);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to terminate app-created Word process PID {ProcessId}", processId);
                    }
                }

                // Cleanup app-created Excel processes
                var excelProcessesToCleanup = _appCreatedExcelProcesses.ToList();
                foreach (var processId in excelProcessesToCleanup)
                {
                    try
                    {
                        var process = Process.GetProcessById(processId);
                        if (!process.HasExited)
                        {
                            _logger.Information("Terminating app-created Excel process PID {ProcessId}", processId);
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                        _appCreatedExcelProcesses.Remove(processId);
                        process.Dispose();
                    }
                    catch (ArgumentException)
                    {
                        // Process already gone
                        _appCreatedExcelProcesses.Remove(processId);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to terminate app-created Excel process PID {ProcessId}", processId);
                    }
                }

                _logger.Information("Cleanup completed. Terminated {WordCount} Word and {ExcelCount} Excel app-created processes",
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