using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Win32;
using Serilog;
using Microsoft.Extensions.DependencyInjection;
using DocHandler.Services;

namespace DocHandler.Services
{
    public class ApplicationHealthChecker : IDisposable
    {
        private static readonly ILogger _logger = Log.ForContext<ApplicationHealthChecker>();
        private readonly Timer _healthCheckTimer;
        private readonly PerformanceMonitor _performanceMonitor;
        private readonly object _healthLock = new object();
        
        // Health status tracking
        private HealthStatus _currentStatus = HealthStatus.Unknown;
        private List<HealthIssue> _currentIssues = new List<HealthIssue>();
        private DateTime _lastHealthCheck = DateTime.MinValue;
        
        // System requirements
        private const long MIN_AVAILABLE_MEMORY_MB = 100;
        private const long MIN_DISK_SPACE_MB = 500;
        private const int MAX_CPU_USAGE_PERCENT = 90;
        private const int HEALTH_CHECK_INTERVAL_MINUTES = 5;

        public event EventHandler<HealthStatusChangedEventArgs> HealthStatusChanged;

        public ApplicationHealthChecker(PerformanceMonitor performanceMonitor = null)
        {
            _performanceMonitor = performanceMonitor ?? new PerformanceMonitor();
            
            // Start health check timer
            _healthCheckTimer = new Timer(PerformHealthCheck, null, 
                TimeSpan.Zero, TimeSpan.FromMinutes(HEALTH_CHECK_INTERVAL_MINUTES));
            
            _logger.Information("Application health checker initialized");
        }

        /// <summary>
        /// Performs a comprehensive health check of the application
        /// </summary>
        public async Task<HealthCheckResult> PerformHealthCheckAsync()
        {
            var result = new HealthCheckResult
            {
                CheckTime = DateTime.UtcNow,
                Issues = new List<HealthIssue>()
            };

            try
            {
                // Memory health check
                await CheckMemoryHealthAsync(result);
                
                // Disk space health check
                await CheckDiskSpaceHealthAsync(result);
                
                // CPU health check
                await CheckCpuHealthAsync(result);
                
                // Office availability check
                await CheckOfficeAvailabilityAsync(result);
                
                // File system permissions check
                await CheckFileSystemPermissionsAsync(result);
                
                // Configuration health check
                await CheckConfigurationHealthAsync(result);
                
                // Service dependencies check
                await CheckServiceDependenciesAsync(result);
                
                // Performance metrics check
                await CheckPerformanceMetricsAsync(result);

                // Determine overall health status
                result.Status = DetermineHealthStatus(result.Issues);
                
                // Update current status
                lock (_healthLock)
                {
                    var previousStatus = _currentStatus;
                    _currentStatus = result.Status;
                    _currentIssues = result.Issues.ToList();
                    _lastHealthCheck = result.CheckTime;
                    
                    if (previousStatus != _currentStatus)
                    {
                        HealthStatusChanged?.Invoke(this, new HealthStatusChangedEventArgs
                        {
                            PreviousStatus = previousStatus,
                            CurrentStatus = _currentStatus,
                            Issues = _currentIssues
                        });
                    }
                }
                
                _logger.Information("Health check completed: {Status} with {IssueCount} issues", 
                    result.Status, result.Issues.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Health check failed");
                result.Status = HealthStatus.Critical;
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Critical,
                    Component = "HealthChecker",
                    Message = $"Health check failed: {ex.Message}",
                    Details = ex.ToString()
                });
            }

            return result;
        }

        /// <summary>
        /// Gets the current health status
        /// </summary>
        public HealthStatus GetCurrentStatus()
        {
            lock (_healthLock)
            {
                return _currentStatus;
            }
        }

        /// <summary>
        /// Gets current health issues
        /// </summary>
        public List<HealthIssue> GetCurrentIssues()
        {
            lock (_healthLock)
            {
                return _currentIssues.ToList();
            }
        }

        /// <summary>
        /// Gets a health summary report
        /// </summary>
        public string GetHealthSummary()
        {
            lock (_healthLock)
            {
                var summary = $"Application Health Status: {_currentStatus}\n";
                summary += $"Last Check: {_lastHealthCheck:yyyy-MM-dd HH:mm:ss} UTC\n";
                
                if (_currentIssues.Any())
                {
                    summary += $"Issues ({_currentIssues.Count}):\n";
                    foreach (var issue in _currentIssues.OrderByDescending(i => i.Severity))
                    {
                        summary += $"  [{issue.Severity}] {issue.Component}: {issue.Message}\n";
                    }
                }
                else
                {
                    summary += "No issues detected\n";
                }

                return summary;
            }
        }

        private async Task CheckMemoryHealthAsync(HealthCheckResult result)
        {
            try
            {
                var memoryInfo = _performanceMonitor.GetMemoryInfo();
                var systemInfo = _performanceMonitor.GetSystemPerformanceInfo();

                // Check for high memory usage
                if (memoryInfo.CurrentMemoryMB > 1000) // 1GB threshold
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Warning,
                        Component = "Memory",
                        Message = $"High memory usage: {memoryInfo.CurrentMemoryMB} MB",
                        Details = $"Current: {memoryInfo.CurrentMemoryMB} MB, Peak: {memoryInfo.PeakMemoryMB} MB"
                    });
                }

                // Check for potential memory leaks
                if (memoryInfo.MemoryGrowthMB > 200) // 200MB growth threshold
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Warning,
                        Component = "Memory",
                        Message = $"Potential memory leak: {memoryInfo.MemoryGrowthMB} MB growth since startup",
                        Details = $"Initial: {memoryInfo.InitialMemoryMB} MB, Current: {memoryInfo.CurrentMemoryMB} MB"
                    });
                }

                // Check available system memory
                if (systemInfo.AvailableMemoryMB < MIN_AVAILABLE_MEMORY_MB)
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Critical,
                        Component = "System Memory",
                        Message = $"Low system memory: {systemInfo.AvailableMemoryMB} MB available",
                        Details = $"Available: {systemInfo.AvailableMemoryMB} MB, Minimum required: {MIN_AVAILABLE_MEMORY_MB} MB"
                    });
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "Memory Check",
                    Message = "Failed to check memory health",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckDiskSpaceHealthAsync(HealthCheckResult result)
        {
            try
            {
                var drives = DriveInfo.GetDrives().Where(d => d.DriveType == DriveType.Fixed && d.IsReady);
                
                foreach (var drive in drives)
                {
                    var availableSpaceMB = drive.AvailableFreeSpace / 1024 / 1024;
                    
                    if (availableSpaceMB < MIN_DISK_SPACE_MB)
                    {
                        result.Issues.Add(new HealthIssue
                        {
                            Severity = HealthIssueSeverity.Critical,
                            Component = "Disk Space",
                            Message = $"Low disk space on drive {drive.Name}: {availableSpaceMB} MB available",
                            Details = $"Available: {availableSpaceMB} MB, Minimum required: {MIN_DISK_SPACE_MB} MB"
                        });
                    }
                    else if (availableSpaceMB < MIN_DISK_SPACE_MB * 2)
                    {
                        result.Issues.Add(new HealthIssue
                        {
                            Severity = HealthIssueSeverity.Warning,
                            Component = "Disk Space",
                            Message = $"Low disk space on drive {drive.Name}: {availableSpaceMB} MB available",
                            Details = $"Available: {availableSpaceMB} MB, Recommended minimum: {MIN_DISK_SPACE_MB * 2} MB"
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "Disk Space Check",
                    Message = "Failed to check disk space",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckCpuHealthAsync(HealthCheckResult result)
        {
            try
            {
                var systemInfo = _performanceMonitor.GetSystemPerformanceInfo();
                
                if (systemInfo.CpuUsagePercent > MAX_CPU_USAGE_PERCENT)
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Warning,
                        Component = "CPU",
                        Message = $"High CPU usage: {systemInfo.CpuUsagePercent:F1}%",
                        Details = $"Current: {systemInfo.CpuUsagePercent:F1}%, Threshold: {MAX_CPU_USAGE_PERCENT}%"
                    });
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "CPU Check",
                    Message = "Failed to check CPU usage",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckOfficeAvailabilityAsync(HealthCheckResult result)
        {
            try
            {
                // Check if Office is installed
                var officeInstalled = IsOfficeInstalled();
                if (!officeInstalled)
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Warning,
                        Component = "Microsoft Office",
                        Message = "Microsoft Office not detected",
                        Details = "Office file conversion features may not work properly"
                    });
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "Office Check",
                    Message = "Failed to check Office availability",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckFileSystemPermissionsAsync(HealthCheckResult result)
        {
            try
            {
                // Check write permissions to temp directory
                var tempDir = Path.GetTempPath();
                var testFile = Path.Combine(tempDir, $"DocHandler_Health_Test_{Guid.NewGuid()}.tmp");
                
                try
                {
                    File.WriteAllText(testFile, "test");
                    File.Delete(testFile);
                }
                catch (Exception ex)
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Critical,
                        Component = "File System",
                        Message = "Cannot write to temporary directory",
                        Details = $"Temp directory: {tempDir}, Error: {ex.Message}"
                    });
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "File System Check",
                    Message = "Failed to check file system permissions",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckConfigurationHealthAsync(HealthCheckResult result)
        {
            try
            {
                // Check if configuration service is accessible - use DI
                var services = new ServiceCollection();
                services.RegisterServices();
                using var serviceProvider = services.BuildServiceProvider();
                var configService = serviceProvider.GetRequiredService<IConfigurationService>();
                var config = configService.Config;
                
                if (string.IsNullOrEmpty(config.DefaultSaveLocation))
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Info,
                        Component = "Configuration",
                        Message = "Default save location not configured",
                        Details = "Users will need to select a save location for each operation"
                    });
                }
                else if (!Directory.Exists(config.DefaultSaveLocation))
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Warning,
                        Component = "Configuration",
                        Message = "Default save location does not exist",
                        Details = $"Path: {config.DefaultSaveLocation}"
                    });
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "Configuration Check",
                    Message = "Failed to check configuration health",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckServiceDependenciesAsync(HealthCheckResult result)
        {
            try
            {
                // Check if required services can be instantiated
                var services = new[]
                {
                    typeof(CompanyNameService),
                    typeof(ScopeOfWorkService),
                    typeof(PdfOperationsService)
                };

                foreach (var serviceType in services)
                {
                    try
                    {
                        var service = Activator.CreateInstance(serviceType);
                        if (service is IDisposable disposable)
                        {
                            disposable.Dispose();
                        }
                    }
                    catch (Exception ex)
                    {
                        result.Issues.Add(new HealthIssue
                        {
                            Severity = HealthIssueSeverity.Critical,
                            Component = "Service Dependencies",
                            Message = $"Failed to initialize {serviceType.Name}",
                            Details = ex.Message
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "Service Dependencies Check",
                    Message = "Failed to check service dependencies",
                    Details = ex.Message
                });
            }
        }

        private async Task CheckPerformanceMetricsAsync(HealthCheckResult result)
        {
            try
            {
                var performanceIssues = _performanceMonitor.CheckPerformanceIssues();
                
                foreach (var issue in performanceIssues)
                {
                    result.Issues.Add(new HealthIssue
                    {
                        Severity = HealthIssueSeverity.Warning,
                        Component = "Performance",
                        Message = issue,
                        Details = "Performance monitoring detected potential issues"
                    });
                }
            }
            catch (Exception ex)
            {
                result.Issues.Add(new HealthIssue
                {
                    Severity = HealthIssueSeverity.Warning,
                    Component = "Performance Metrics Check",
                    Message = "Failed to check performance metrics",
                    Details = ex.Message
                });
            }
        }

        private bool IsOfficeInstalled()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office"))
                {
                    return key != null && key.GetSubKeyNames().Any(name => name.Contains("Word") || name.Contains("Excel"));
                }
            }
            catch
            {
                return false;
            }
        }

        private HealthStatus DetermineHealthStatus(List<HealthIssue> issues)
        {
            if (!issues.Any())
                return HealthStatus.Healthy;

            if (issues.Any(i => i.Severity == HealthIssueSeverity.Critical))
                return HealthStatus.Critical;

            if (issues.Any(i => i.Severity == HealthIssueSeverity.Warning))
                return HealthStatus.Warning;

            return HealthStatus.Info;
        }

        private void PerformHealthCheck(object state)
        {
            try
            {
                _ = Task.Run(async () => await PerformHealthCheckAsync());
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Health check timer failed");
            }
        }

        public void Dispose()
        {
            _healthCheckTimer?.Dispose();
            _performanceMonitor?.Dispose();
            _logger.Information("Application health checker disposed");
        }
    }

    public enum HealthStatus
    {
        Unknown,
        Healthy,
        Info,
        Warning,
        Critical
    }

    public enum HealthIssueSeverity
    {
        Info,
        Warning,
        Critical
    }

    public class HealthIssue
    {
        public HealthIssueSeverity Severity { get; set; }
        public string Component { get; set; } = "";
        public string Message { get; set; } = "";
        public string Details { get; set; } = "";
        public DateTime DetectedAt { get; set; } = DateTime.UtcNow;
    }

    public class HealthCheckResult
    {
        public DateTime CheckTime { get; set; }
        public HealthStatus Status { get; set; }
        public List<HealthIssue> Issues { get; set; } = new List<HealthIssue>();
    }

    public class HealthStatusChangedEventArgs : EventArgs
    {
        public HealthStatus PreviousStatus { get; set; }
        public HealthStatus CurrentStatus { get; set; }
        public List<HealthIssue> Issues { get; set; } = new List<HealthIssue>();
    }
} 