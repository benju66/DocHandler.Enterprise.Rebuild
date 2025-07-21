using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Linq;
using System.Text;
using System.IO;
using Serilog;

namespace DocHandler.Services
{
    public class PerformanceMonitor : IDisposable
    {
        private static readonly ILogger _logger = Log.ForContext<PerformanceMonitor>();
        private readonly Timer _memoryTimer;
        private readonly PerformanceCounter _cpuCounter;
        private readonly PerformanceCounter _ramCounter;
        private readonly Process _currentProcess;
        private readonly object _metricsLock = new object();

        // Performance metrics
        private readonly Dictionary<string, List<double>> _metrics = new Dictionary<string, List<double>>();
        private readonly Dictionary<string, DateTime> _operationStartTimes = new Dictionary<string, DateTime>();
        
        // Document processing statistics
        private readonly Dictionary<string, DocumentTypeStats> _documentProcessingStats = new Dictionary<string, DocumentTypeStats>();
        private readonly object _docStatsLock = new object();
        
        // Memory tracking
        private long _initialMemoryUsage;
        private long _peakMemoryUsage;
        private long _currentMemoryUsage;
        
        // Performance thresholds
        private const int HIGH_MEMORY_THRESHOLD_MB = 500;
        private const int MEMORY_LEAK_THRESHOLD_MB = 300; // Increased from 100MB to reduce false positives
        private const int MAX_METRIC_HISTORY = 100;
        
        // Memory pressure monitoring
        private long _memoryThresholdBytes;
        private Timer _memoryMonitorTimer;
        
        public event EventHandler<MemoryPressureEventArgs>? MemoryPressureDetected;

        public PerformanceMonitor(int memoryLimitMB = 500)
        {
            _currentProcess = Process.GetCurrentProcess();
            _initialMemoryUsage = _currentProcess.WorkingSet64;
            _peakMemoryUsage = _initialMemoryUsage;
            _memoryThresholdBytes = memoryLimitMB * 1024L * 1024L;
            
            try
            {
                _cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                _ramCounter = new PerformanceCounter("Memory", "Available MBytes");
                _cpuCounter.NextValue(); // Initialize
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to initialize performance counters");
            }

            // Start memory monitoring timer (every 30 seconds)
            _memoryTimer = new Timer(MonitorMemory, null, TimeSpan.Zero, TimeSpan.FromSeconds(30));
            
            // Start memory pressure monitoring
            _memoryMonitorTimer = new Timer(CheckMemoryPressure, null, 
                5000, 5000); // 5 seconds in milliseconds
            
            _logger.Information("Performance monitor initialized with {InitialMemory} MB initial memory", 
                _initialMemoryUsage / 1024 / 1024);
        }

        /// <summary>
        /// Starts tracking performance for an operation
        /// </summary>
        public void StartOperation(string operationName)
        {
            lock (_metricsLock)
            {
                _operationStartTimes[operationName] = DateTime.UtcNow;
            }
        }

        /// <summary>
        /// Ends tracking and records performance metrics for an operation
        /// </summary>
        public void EndOperation(string operationName)
        {
            lock (_metricsLock)
            {
                if (_operationStartTimes.TryGetValue(operationName, out var startTime))
                {
                    var duration = DateTime.UtcNow - startTime;
                    RecordMetric($"{operationName}_Duration", duration.TotalMilliseconds);
                    _operationStartTimes.Remove(operationName);
                }
            }
        }

        /// <summary>
        /// Records a custom metric
        /// </summary>
        public void RecordMetric(string metricName, double value)
        {
            lock (_metricsLock)
            {
                if (!_metrics.ContainsKey(metricName))
                {
                    _metrics[metricName] = new List<double>();
                }

                _metrics[metricName].Add(value);
                
                // Keep only the last MAX_METRIC_HISTORY values
                if (_metrics[metricName].Count > MAX_METRIC_HISTORY)
                {
                    _metrics[metricName].RemoveAt(0);
                }
            }
        }

        /// <summary>
        /// Gets performance statistics for a metric
        /// </summary>
        public PerformanceStats GetMetricStats(string metricName)
        {
            lock (_metricsLock)
            {
                if (!_metrics.ContainsKey(metricName) || _metrics[metricName].Count == 0)
                {
                    return new PerformanceStats();
                }

                var values = _metrics[metricName];
                return new PerformanceStats
                {
                    Count = values.Count,
                    Average = values.Average(),
                    Min = values.Min(),
                    Max = values.Max(),
                    Latest = values.Last()
                };
            }
        }

        /// <summary>
        /// Gets current memory usage information
        /// </summary>
        public MemoryInfo GetMemoryInfo()
        {
            _currentProcess.Refresh();
            var currentMemory = _currentProcess.WorkingSet64;
            
            if (currentMemory > _peakMemoryUsage)
            {
                _peakMemoryUsage = currentMemory;
            }

            return new MemoryInfo
            {
                InitialMemoryMB = _initialMemoryUsage / 1024 / 1024,
                CurrentMemoryMB = currentMemory / 1024 / 1024,
                PeakMemoryMB = _peakMemoryUsage / 1024 / 1024,
                MemoryGrowthMB = (currentMemory - _initialMemoryUsage) / 1024 / 1024,
                GCTotalMemoryMB = GC.GetTotalMemory(false) / 1024 / 1024,
                Gen0Collections = GC.CollectionCount(0),
                Gen1Collections = GC.CollectionCount(1),
                Gen2Collections = GC.CollectionCount(2)
            };
        }

        /// <summary>
        /// Gets current system performance information
        /// </summary>
        public SystemPerformanceInfo GetSystemPerformanceInfo()
        {
            var info = new SystemPerformanceInfo();
            
            try
            {
                if (_cpuCounter != null)
                {
                    info.CpuUsagePercent = _cpuCounter.NextValue();
                }
                
                if (_ramCounter != null)
                {
                    info.AvailableMemoryMB = (long)_ramCounter.NextValue();
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to get system performance counters");
            }

            return info;
        }

        /// <summary>
        /// Gets a comprehensive performance summary
        /// </summary>
        public string GetPerformanceSummary()
        {
            var memoryInfo = GetMemoryInfo();
            var systemInfo = GetSystemPerformanceInfo();
            
            var summary = $"Performance Summary:\n";
            summary += $"  Memory Usage: {memoryInfo.CurrentMemoryMB:F1} MB (Peak: {memoryInfo.PeakMemoryMB:F1} MB)\n";
            summary += $"  Memory Growth: {memoryInfo.MemoryGrowthMB:F1} MB from startup\n";
            summary += $"  GC Collections: Gen0={memoryInfo.Gen0Collections}, Gen1={memoryInfo.Gen1Collections}, Gen2={memoryInfo.Gen2Collections}\n";
            summary += $"  System CPU: {systemInfo.CpuUsagePercent:F1}%\n";
            summary += $"  Available RAM: {systemInfo.AvailableMemoryMB} MB\n";
            
            lock (_metricsLock)
            {
                if (_metrics.Any())
                {
                    summary += "  Operation Metrics:\n";
                    foreach (var metric in _metrics.Keys.Take(10)) // Show top 10 metrics
                    {
                        var stats = GetMetricStats(metric);
                        summary += $"    {metric}: Avg={stats.Average:F1}ms, Count={stats.Count}\n";
                    }
                }
            }

            return summary;
        }

        /// <summary>
        /// Checks if there are any performance issues
        /// </summary>
        public List<string> CheckPerformanceIssues()
        {
            var issues = new List<string>();
            var memoryInfo = GetMemoryInfo();
            
            // Check for high memory usage
            if (memoryInfo.CurrentMemoryMB > HIGH_MEMORY_THRESHOLD_MB)
            {
                issues.Add($"High memory usage: {memoryInfo.CurrentMemoryMB:F1} MB (threshold: {HIGH_MEMORY_THRESHOLD_MB} MB)");
            }

            // Check for potential memory leaks
            if (memoryInfo.MemoryGrowthMB > MEMORY_LEAK_THRESHOLD_MB)
            {
                issues.Add($"Potential memory leak: {memoryInfo.MemoryGrowthMB:F1} MB growth since startup");
            }

            // Check for excessive GC collections
            if (memoryInfo.Gen2Collections > 100)
            {
                issues.Add($"Excessive Gen2 garbage collections: {memoryInfo.Gen2Collections}");
            }

            // Check for slow operations
            lock (_metricsLock)
            {
                foreach (var metric in _metrics.Keys.Where(k => k.Contains("_Duration")))
                {
                    var stats = GetMetricStats(metric);
                    if (stats.Average > 5000) // Operations taking more than 5 seconds on average
                    {
                        issues.Add($"Slow operation detected: {metric} averaging {stats.Average:F1}ms");
                    }
                }
            }

            return issues;
        }

        /// <summary>
        /// Forces garbage collection and logs memory stats
        /// </summary>
        public void ForceGarbageCollection()
        {
            var beforeMemory = GC.GetTotalMemory(false);
            
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            
            var afterMemory = GC.GetTotalMemory(false);
            var freedMemory = (beforeMemory - afterMemory) / 1024 / 1024;
            
            _logger.Information("Forced garbage collection freed {FreedMemory} MB", freedMemory);
        }

        private void MonitorMemory(object state)
        {
            try
            {
                var memoryInfo = GetMemoryInfo();
                _currentMemoryUsage = memoryInfo.CurrentMemoryMB * 1024 * 1024;
                
                // Check for potential issues
                var issues = CheckPerformanceIssues();
                if (issues.Any())
                {
                    _logger.Warning("Performance issues detected: {Issues}", string.Join(", ", issues));
                }
                
                // Log memory status periodically
                if (DateTime.UtcNow.Second % 300 == 0) // Every 5 minutes
                {
                    _logger.Information("Memory status: Current={CurrentMemory}MB, Peak={PeakMemory}MB, Growth={Growth}MB", 
                        memoryInfo.CurrentMemoryMB, memoryInfo.PeakMemoryMB, memoryInfo.MemoryGrowthMB);
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during memory monitoring");
            }
        }
        
        private void CheckMemoryPressure(object? state)
        {
            var process = Process.GetCurrentProcess();
            var workingSet = process.WorkingSet64;
            
            if (workingSet > _memoryThresholdBytes)
            {
                var args = new MemoryPressureEventArgs
                {
                    CurrentMemoryMB = workingSet / (1024 * 1024),
                    ThresholdMB = _memoryThresholdBytes / (1024 * 1024),
                    IsCritical = workingSet > _memoryThresholdBytes * 1.5
                };
                
                MemoryPressureDetected?.Invoke(this, args);
                
                // Force garbage collection if critical
                if (args.IsCritical)
                {
                    _logger.Warning("Critical memory pressure detected: {CurrentMB}MB", args.CurrentMemoryMB);
                    
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }
        }

        /// <summary>
        /// Records document processing metrics with enhanced details
        /// </summary>
        public void RecordDocumentProcessed(string filePath, double processingTimeMs, bool success, 
            long fileSize = 0, string errorType = null)
        {
            var extension = Path.GetExtension(filePath)?.ToLowerInvariant() ?? "unknown";
            var fileName = Path.GetFileName(filePath);
            
            // Map extensions to readable types
            string documentType = extension switch
            {
                ".pdf" => "PDF",
                ".docx" => "Word (Modern)",
                ".doc" => "Word (Legacy)",
                ".xlsx" => "Excel (Modern)",
                ".xls" => "Excel (Legacy)",
                ".txt" => "Text",
                _ => $"Other ({extension})"
            };
            
            lock (_docStatsLock)
            {
                if (!_documentProcessingStats.ContainsKey(documentType))
                {
                    _documentProcessingStats[documentType] = new DocumentTypeStats();
                }
                
                var stats = _documentProcessingStats[documentType];
                stats.TotalFiles++;
                if (success) 
                {
                    stats.SuccessfulFiles++;
                }
                else if (!string.IsNullOrEmpty(errorType))
                {
                    if (!stats.ErrorTypes.ContainsKey(errorType))
                        stats.ErrorTypes[errorType] = 0;
                    stats.ErrorTypes[errorType]++;
                }
                
                stats.TotalProcessingTime += processingTimeMs;
                stats.MinTime = Math.Min(stats.MinTime, processingTimeMs);
                stats.MaxTime = Math.Max(stats.MaxTime, processingTimeMs);
                stats.LastProcessed = DateTime.Now;
                stats.TotalBytesProcessed += fileSize;
                
                // Keep recent files
                if (stats.RecentFiles.Count >= 5)
                    stats.RecentFiles.Dequeue();
                stats.RecentFiles.Enqueue(fileName);
            }
            
            // Also record in existing metrics
            RecordMetric($"DocProcess_{extension}", processingTimeMs);
        }

        /// <summary>
        /// Gets document processing performance summary
        /// </summary>
        public string GetDocumentProcessingPerformanceSummary()
        {
            lock (_docStatsLock)
            {
                if (!_documentProcessingStats.Any())
                    return "Document Processing Performance:\nNo document processing metrics available yet.";
                
                var totalFiles = _documentProcessingStats.Sum(kvp => kvp.Value.TotalFiles);
                var totalTime = _documentProcessingStats.Sum(kvp => kvp.Value.TotalProcessingTime);
                var totalSuccess = _documentProcessingStats.Sum(kvp => kvp.Value.SuccessfulFiles);
                var totalBytes = _documentProcessingStats.Sum(kvp => kvp.Value.TotalBytesProcessed);
                var overallAvg = totalFiles > 0 ? totalTime / totalFiles : 0;
                var overallSuccessRate = totalFiles > 0 ? (double)totalSuccess / totalFiles * 100 : 0;
                var overallThroughput = totalBytes > 0 && totalTime > 0 
                    ? (totalBytes / 1024.0 / 1024.0) / (totalTime / 1000.0) 
                    : 0;
                
                var summary = new StringBuilder();
                summary.AppendLine("Document Processing Performance:");
                summary.AppendLine($"Total: {totalFiles} docs, Overall Avg: {overallAvg:F1}ms, Success Rate: {overallSuccessRate:F1}%, Throughput: {overallThroughput:F2} MB/s");
                summary.AppendLine("  By Type:");
                
                foreach (var kvp in _documentProcessingStats.OrderBy(x => x.Key))
                {
                    var type = kvp.Key;
                    var stats = kvp.Value;
                    summary.AppendLine($"    {type}: {stats.TotalFiles} files, Avg: {stats.AverageTime:F1}ms, Success: {stats.SuccessRate:F1}%, Throughput: {stats.ThroughputMBps:F2} MB/s");
                    
                    // Add error breakdown if any
                    if (stats.ErrorTypes.Any())
                    {
                        summary.AppendLine($"      Errors: {string.Join(", ", stats.ErrorTypes.Select(et => $"{et.Key} ({et.Value})"))}");
                    }
                }
                
                return summary.ToString();
            }
        }

        /// <summary>
        /// Gets detailed document type statistics for advanced UI
        /// </summary>
        public Dictionary<string, DocumentTypeStats> GetDocumentTypeStats()
        {
            lock (_docStatsLock)
            {
                // Return a deep copy to avoid threading issues
                var copy = new Dictionary<string, DocumentTypeStats>();
                foreach (var kvp in _documentProcessingStats)
                {
                    copy[kvp.Key] = new DocumentTypeStats
                    {
                        TotalFiles = kvp.Value.TotalFiles,
                        SuccessfulFiles = kvp.Value.SuccessfulFiles,
                        TotalProcessingTime = kvp.Value.TotalProcessingTime,
                        MinTime = kvp.Value.MinTime,
                        MaxTime = kvp.Value.MaxTime,
                        FirstProcessed = kvp.Value.FirstProcessed,
                        LastProcessed = kvp.Value.LastProcessed,
                        TotalBytesProcessed = kvp.Value.TotalBytesProcessed,
                        ErrorTypes = new Dictionary<string, int>(kvp.Value.ErrorTypes),
                        RecentFiles = new Queue<string>(kvp.Value.RecentFiles)
                    };
                }
                return copy;
            }
        }

        public void Dispose()
        {
            _memoryTimer?.Dispose();
            _memoryMonitorTimer?.Dispose();
            _cpuCounter?.Dispose();
            _ramCounter?.Dispose();
            _currentProcess?.Dispose();
            
            _logger.Information("Performance monitor disposed");
        }
    }

    public class DocumentTypeStats
    {
        public int TotalFiles { get; set; }
        public int SuccessfulFiles { get; set; }
        public double TotalProcessingTime { get; set; }
        public double MinTime { get; set; } = double.MaxValue;
        public double MaxTime { get; set; }
        public DateTime FirstProcessed { get; set; } = DateTime.Now;
        public DateTime LastProcessed { get; set; } = DateTime.Now;
        public long TotalBytesProcessed { get; set; }
        public Dictionary<string, int> ErrorTypes { get; set; } = new Dictionary<string, int>();
        public Queue<string> RecentFiles { get; set; } = new Queue<string>(5);
        
        public double AverageTime => TotalFiles > 0 ? TotalProcessingTime / TotalFiles : 0;
        public double SuccessRate => TotalFiles > 0 ? (double)SuccessfulFiles / TotalFiles * 100 : 0;
        public double ThroughputMBps => TotalBytesProcessed > 0 && TotalProcessingTime > 0 
            ? (TotalBytesProcessed / 1024.0 / 1024.0) / (TotalProcessingTime / 1000.0) 
            : 0;
    }

    public class PerformanceStats
    {
        public int Count { get; set; }
        public double Average { get; set; }
        public double Min { get; set; }
        public double Max { get; set; }
        public double Latest { get; set; }
    }

    public class MemoryInfo
    {
        public long InitialMemoryMB { get; set; }
        public long CurrentMemoryMB { get; set; }
        public long PeakMemoryMB { get; set; }
        public long MemoryGrowthMB { get; set; }
        public long GCTotalMemoryMB { get; set; }
        public int Gen0Collections { get; set; }
        public int Gen1Collections { get; set; }
        public int Gen2Collections { get; set; }
    }

    public class SystemPerformanceInfo
    {
        public float CpuUsagePercent { get; set; }
        public long AvailableMemoryMB { get; set; }
    }
    
    public class MemoryPressureEventArgs : EventArgs
    {
        public long CurrentMemoryMB { get; set; }
        public long ThresholdMB { get; set; }
        public bool IsCritical { get; set; }
    }
} 