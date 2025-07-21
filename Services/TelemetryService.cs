using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Text.Json;
using Serilog;

namespace DocHandler.Services
{
    public class TelemetryService : ITelemetryService
    {
        private static readonly ILogger _logger = Log.ForContext<TelemetryService>();
        private readonly Timer _flushTimer;
        private readonly object _telemetryLock = new object();
        private readonly List<TelemetryEvent> _pendingEvents = new List<TelemetryEvent>();
        private readonly Dictionary<string, object> _sessionMetrics = new Dictionary<string, object>();
        private readonly string _telemetryFilePath;
        private readonly PerformanceMonitor _performanceMonitor;
        
        // Configuration
        private const int MAX_PENDING_EVENTS = 100;
        private const int FLUSH_INTERVAL_MINUTES = 5;
        private const int MAX_TELEMETRY_FILE_SIZE_MB = 10;
        private const int MAX_TELEMETRY_FILES = 5;
        
        // Session tracking
        private readonly Guid _sessionId;
        private readonly DateTime _sessionStartTime;
        private int _totalOperations;
        private int _successfulOperations;
        private int _failedOperations;
        
        public TelemetryService(PerformanceMonitor performanceMonitor = null)
        {
            _performanceMonitor = performanceMonitor ?? new PerformanceMonitor();
            _sessionId = Guid.NewGuid();
            _sessionStartTime = DateTime.UtcNow;
            
            // Setup telemetry file path
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var telemetryDir = Path.Combine(appDataPath, "DocHandler", "Telemetry");
            Directory.CreateDirectory(telemetryDir);
            _telemetryFilePath = Path.Combine(telemetryDir, $"telemetry_{DateTime.Now:yyyyMMdd}.json");
            
            // Setup flush timer
            _flushTimer = new Timer(FlushTelemetry, null, 
                TimeSpan.FromMinutes(FLUSH_INTERVAL_MINUTES), 
                TimeSpan.FromMinutes(FLUSH_INTERVAL_MINUTES));
            
            // Initialize session metrics
            InitializeSessionMetrics();
            
            // Track session start
            TrackEvent("SessionStart", new Dictionary<string, object>
            {
                ["SessionId"] = _sessionId,
                ["StartTime"] = _sessionStartTime,
                ["Version"] = GetApplicationVersion()
            });
            
            _logger.Information("Telemetry service initialized for session {SessionId}", _sessionId);
        }
        
        /// <summary>
        /// Tracks a telemetry event
        /// </summary>
        public void TrackEvent(string eventName, Dictionary<string, object> properties = null, Dictionary<string, double> metrics = null)
        {
            var telemetryEvent = new TelemetryEvent
            {
                Id = Guid.NewGuid(),
                SessionId = _sessionId,
                EventName = eventName,
                Timestamp = DateTime.UtcNow,
                Properties = properties ?? new Dictionary<string, object>(),
                Metrics = metrics ?? new Dictionary<string, double>()
            };
            
            // Add common properties
            telemetryEvent.Properties["SessionId"] = _sessionId;
            telemetryEvent.Properties["MachineName"] = Environment.MachineName;
            telemetryEvent.Properties["OSVersion"] = Environment.OSVersion.ToString();
            telemetryEvent.Properties["ProcessorCount"] = Environment.ProcessorCount;
            
            lock (_telemetryLock)
            {
                _pendingEvents.Add(telemetryEvent);
                
                // Flush if we have too many pending events
                if (_pendingEvents.Count >= MAX_PENDING_EVENTS)
                {
                    FlushTelemetryInternal();
                }
            }
            
            _logger.Debug("Telemetry event tracked: {EventName}", eventName);
        }
        
        /// <summary>
        /// Tracks an operation with performance metrics
        /// </summary>
        public void TrackOperation(string operationName, TimeSpan duration, bool successful, string details = null)
        {
            Interlocked.Increment(ref _totalOperations);
            
            if (successful)
            {
                Interlocked.Increment(ref _successfulOperations);
            }
            else
            {
                Interlocked.Increment(ref _failedOperations);
            }
            
            var properties = new Dictionary<string, object>
            {
                ["OperationName"] = operationName,
                ["Successful"] = successful,
                ["Duration"] = duration.TotalMilliseconds
            };
            
            if (!string.IsNullOrEmpty(details))
            {
                properties["Details"] = details;
            }
            
            var metrics = new Dictionary<string, double>
            {
                ["DurationMs"] = duration.TotalMilliseconds,
                ["Success"] = successful ? 1.0 : 0.0
            };
            
            TrackEvent("OperationCompleted", properties, metrics);
            
            // Update session metrics
            UpdateSessionMetrics(operationName, duration, successful);
        }
        
        /// <summary>
        /// Tracks an exception
        /// </summary>
        public void TrackException(Exception exception, string context = null)
        {
            var properties = new Dictionary<string, object>
            {
                ["ExceptionType"] = exception.GetType().Name,
                ["Message"] = exception.Message,
                ["StackTrace"] = exception.StackTrace
            };
            
            if (!string.IsNullOrEmpty(context))
            {
                properties["Context"] = context;
            }
            
            if (exception.InnerException != null)
            {
                properties["InnerExceptionType"] = exception.InnerException.GetType().Name;
                properties["InnerExceptionMessage"] = exception.InnerException.Message;
            }
            
            TrackEvent("ExceptionOccurred", properties);
            
            _logger.Warning(exception, "Exception tracked in telemetry: {Context}", context);
        }
        
        /// <summary>
        /// Tracks user interaction
        /// </summary>
        public void TrackUserInteraction(string action, Dictionary<string, object> properties = null)
        {
            var eventProperties = new Dictionary<string, object>
            {
                ["Action"] = action,
                ["Timestamp"] = DateTime.UtcNow
            };
            
            if (properties != null)
            {
                foreach (var prop in properties)
                {
                    eventProperties[prop.Key] = prop.Value;
                }
            }
            
            TrackEvent("UserInteraction", eventProperties);
        }
        
        /// <summary>
        /// Tracks file processing statistics
        /// </summary>
        public void TrackFileProcessing(string fileType, long fileSize, TimeSpan processingTime, bool successful)
        {
            var properties = new Dictionary<string, object>
            {
                ["FileType"] = fileType,
                ["FileSize"] = fileSize,
                ["Successful"] = successful
            };
            
            var metrics = new Dictionary<string, double>
            {
                ["FileSizeBytes"] = fileSize,
                ["ProcessingTimeMs"] = processingTime.TotalMilliseconds,
                ["Success"] = successful ? 1.0 : 0.0
            };
            
            TrackEvent("FileProcessed", properties, metrics);
        }
        
        /// <summary>
        /// Gets current session metrics
        /// </summary>
        public Dictionary<string, object> GetSessionMetrics()
        {
            lock (_telemetryLock)
            {
                var metrics = new Dictionary<string, object>(_sessionMetrics)
                {
                    ["SessionId"] = _sessionId,
                    ["SessionDuration"] = DateTime.UtcNow - _sessionStartTime,
                    ["TotalOperations"] = _totalOperations,
                    ["SuccessfulOperations"] = _successfulOperations,
                    ["FailedOperations"] = _failedOperations,
                    ["SuccessRate"] = _totalOperations > 0 ? (double)_successfulOperations / _totalOperations : 0.0
                };
                
                // Add performance metrics
                var memoryInfo = _performanceMonitor.GetMemoryInfo();
                metrics["CurrentMemoryMB"] = memoryInfo.CurrentMemoryMB;
                metrics["PeakMemoryMB"] = memoryInfo.PeakMemoryMB;
                metrics["MemoryGrowthMB"] = memoryInfo.MemoryGrowthMB;
                
                return metrics;
            }
        }
        
        /// <summary>
        /// Gets telemetry summary
        /// </summary>
        public string GetTelemetrySummary()
        {
            var metrics = GetSessionMetrics();
            var summary = $"Telemetry Summary (Session: {_sessionId})\n";
            summary += $"  Session Duration: {metrics["SessionDuration"]}\n";
            summary += $"  Total Operations: {metrics["TotalOperations"]}\n";
            summary += $"  Success Rate: {metrics["SuccessRate"]:P2}\n";
            summary += $"  Memory Usage: {metrics["CurrentMemoryMB"]} MB (Peak: {metrics["PeakMemoryMB"]} MB)\n";
            summary += $"  Pending Events: {_pendingEvents.Count}\n";
            
            return summary;
        }
        
        /// <summary>
        /// Flushes telemetry data to file
        /// </summary>
        public void FlushTelemetry()
        {
            lock (_telemetryLock)
            {
                FlushTelemetryInternal();
            }
        }
        
        /// <summary>
        /// Gets telemetry statistics from stored files
        /// </summary>
        public async Task<TelemetryStatistics> GetTelemetryStatisticsAsync()
        {
            var stats = new TelemetryStatistics();
            
            try
            {
                var telemetryDir = Path.GetDirectoryName(_telemetryFilePath);
                var telemetryFiles = Directory.GetFiles(telemetryDir, "telemetry_*.json");
                
                foreach (var file in telemetryFiles.Take(10)) // Limit to recent files
                {
                    try
                    {
                        var content = await File.ReadAllTextAsync(file);
                        var events = JsonSerializer.Deserialize<List<TelemetryEvent>>(content);
                        
                        if (events != null)
                        {
                            stats.TotalEvents += events.Count;
                            stats.TotalSessions += events.Where(e => e.EventName == "SessionStart").Count();
                            stats.TotalOperations += events.Where(e => e.EventName == "OperationCompleted").Count();
                            stats.TotalExceptions += events.Where(e => e.EventName == "ExceptionOccurred").Count();
                            
                            var firstEvent = events.OrderBy(e => e.Timestamp).FirstOrDefault();
                            var lastEvent = events.OrderByDescending(e => e.Timestamp).FirstOrDefault();
                            
                            if (firstEvent != null && (stats.EarliestEvent == DateTime.MinValue || firstEvent.Timestamp < stats.EarliestEvent))
                            {
                                stats.EarliestEvent = firstEvent.Timestamp;
                            }
                            
                            if (lastEvent != null && lastEvent.Timestamp > stats.LatestEvent)
                            {
                                stats.LatestEvent = lastEvent.Timestamp;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to read telemetry file: {File}", file);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to get telemetry statistics");
            }
            
            return stats;
        }
        
        private void InitializeSessionMetrics()
        {
            lock (_telemetryLock)
            {
                _sessionMetrics["SessionStartTime"] = _sessionStartTime;
                _sessionMetrics["ApplicationVersion"] = GetApplicationVersion();
                _sessionMetrics["MachineName"] = Environment.MachineName;
                _sessionMetrics["OSVersion"] = Environment.OSVersion.ToString();
                _sessionMetrics["ProcessorCount"] = Environment.ProcessorCount;
                _sessionMetrics["WorkingSet"] = Environment.WorkingSet;
            }
        }
        
        private void UpdateSessionMetrics(string operationName, TimeSpan duration, bool successful)
        {
            lock (_telemetryLock)
            {
                // Track operation counts by type
                var operationCountKey = $"Operation_{operationName}_Count";
                _sessionMetrics[operationCountKey] = (_sessionMetrics.ContainsKey(operationCountKey) 
                    ? (int)_sessionMetrics[operationCountKey] : 0) + 1;
                
                // Track average duration
                var durationKey = $"Operation_{operationName}_AvgDuration";
                var currentCount = (int)_sessionMetrics[operationCountKey];
                var currentAvg = _sessionMetrics.ContainsKey(durationKey) ? (double)_sessionMetrics[durationKey] : 0.0;
                _sessionMetrics[durationKey] = ((currentAvg * (currentCount - 1)) + duration.TotalMilliseconds) / currentCount;
                
                // Track success rate
                var successKey = $"Operation_{operationName}_SuccessRate";
                var successCount = _sessionMetrics.ContainsKey($"Operation_{operationName}_SuccessCount") 
                    ? (int)_sessionMetrics[$"Operation_{operationName}_SuccessCount"] : 0;
                
                if (successful)
                {
                    successCount++;
                    _sessionMetrics[$"Operation_{operationName}_SuccessCount"] = successCount;
                }
                
                _sessionMetrics[successKey] = (double)successCount / currentCount;
            }
        }
        
        private void FlushTelemetryInternal()
        {
            if (_pendingEvents.Count == 0)
                return;
            
            try
            {
                // Rotate files if needed
                RotateTelemetryFiles();
                
                // Read existing events
                var existingEvents = new List<TelemetryEvent>();
                if (File.Exists(_telemetryFilePath))
                {
                    var content = File.ReadAllText(_telemetryFilePath);
                    if (!string.IsNullOrEmpty(content))
                    {
                        existingEvents = JsonSerializer.Deserialize<List<TelemetryEvent>>(content) ?? new List<TelemetryEvent>();
                    }
                }
                
                // Add pending events
                existingEvents.AddRange(_pendingEvents);
                
                // Serialize and write
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };
                
                var json = JsonSerializer.Serialize(existingEvents, options);
                File.WriteAllText(_telemetryFilePath, json);
                
                _logger.Information("Flushed {Count} telemetry events to file", _pendingEvents.Count);
                _pendingEvents.Clear();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to flush telemetry data");
            }
        }
        
        private void RotateTelemetryFiles()
        {
            try
            {
                var telemetryDir = Path.GetDirectoryName(_telemetryFilePath);
                var telemetryFiles = Directory.GetFiles(telemetryDir, "telemetry_*.json")
                    .OrderByDescending(f => File.GetCreationTime(f))
                    .ToList();
                
                // Remove old files
                while (telemetryFiles.Count >= MAX_TELEMETRY_FILES)
                {
                    var oldestFile = telemetryFiles.Last();
                    File.Delete(oldestFile);
                    telemetryFiles.RemoveAt(telemetryFiles.Count - 1);
                    _logger.Information("Deleted old telemetry file: {File}", Path.GetFileName(oldestFile));
                }
                
                // Check current file size
                if (File.Exists(_telemetryFilePath))
                {
                    var fileInfo = new FileInfo(_telemetryFilePath);
                    if (fileInfo.Length > MAX_TELEMETRY_FILE_SIZE_MB * 1024 * 1024)
                    {
                        // Archive current file
                        var archivePath = Path.ChangeExtension(_telemetryFilePath, $".{DateTime.Now:HHmmss}.json");
                        File.Move(_telemetryFilePath, archivePath);
                        _logger.Information("Archived telemetry file: {File}", Path.GetFileName(archivePath));
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to rotate telemetry files");
            }
        }
        
        private string GetApplicationVersion()
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                var version = assembly.GetName().Version;
                return version?.ToString() ?? "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }
        
        private void FlushTelemetry(object state)
        {
            try
            {
                FlushTelemetry();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Telemetry flush timer failed");
            }
        }
        
        public void Dispose()
        {
            try
            {
                // Track session end
                TrackEvent("SessionEnd", new Dictionary<string, object>
                {
                    ["SessionId"] = _sessionId,
                    ["SessionDuration"] = DateTime.UtcNow - _sessionStartTime,
                    ["TotalOperations"] = _totalOperations,
                    ["SuccessfulOperations"] = _successfulOperations,
                    ["FailedOperations"] = _failedOperations
                });
                
                // Flush remaining events
                FlushTelemetry();
                
                _flushTimer?.Dispose();
                _performanceMonitor?.Dispose();
                
                _logger.Information("Telemetry service disposed for session {SessionId}", _sessionId);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during telemetry service disposal");
            }
        }
    }
    
    public class TelemetryEvent
    {
        public Guid Id { get; set; }
        public Guid SessionId { get; set; }
        public string EventName { get; set; } = "";
        public DateTime Timestamp { get; set; }
        public Dictionary<string, object> Properties { get; set; } = new Dictionary<string, object>();
        public Dictionary<string, double> Metrics { get; set; } = new Dictionary<string, double>();
    }
    
    public class TelemetryStatistics
    {
        public int TotalEvents { get; set; }
        public int TotalSessions { get; set; }
        public int TotalOperations { get; set; }
        public int TotalExceptions { get; set; }
        public DateTime EarliestEvent { get; set; } = DateTime.MinValue;
        public DateTime LatestEvent { get; set; } = DateTime.MinValue;
    }
} 