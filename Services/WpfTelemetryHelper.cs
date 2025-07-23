using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// WPF-specific telemetry helper that provides easy integration with desktop applications
    /// and automatically collects WPF and desktop-specific metrics
    /// </summary>
    public class WpfTelemetryHelper : IDisposable
    {
        private static readonly ILogger _logger = Log.ForContext<WpfTelemetryHelper>();
        private readonly TelemetryService _telemetryService;
        private readonly PerformanceMonitor _performanceMonitor;
        private readonly Timer _performanceTimer;
        private readonly Timer _memoryTimer;
        
        private bool _disposed = false;
        private readonly DateTime _sessionStart;
        private readonly Dictionary<string, DateTime> _windowOpenTimes = new Dictionary<string, DateTime>();
        private readonly Dictionary<string, int> _uiInteractionCounts = new Dictionary<string, int>();
        
        public WpfTelemetryHelper(TelemetryService telemetryService, PerformanceMonitor performanceMonitor)
        {
            _telemetryService = telemetryService ?? throw new ArgumentNullException(nameof(telemetryService));
            _performanceMonitor = performanceMonitor ?? throw new ArgumentNullException(nameof(performanceMonitor));
            _sessionStart = DateTime.UtcNow;
            
            // Setup performance monitoring timers
            _performanceTimer = new Timer(CollectPerformanceMetrics, null, 
                TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
            _memoryTimer = new Timer(CollectMemoryMetrics, null, 
                TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
            
            _logger.Information("WPF Telemetry Helper initialized");
        }

        /// <summary>
        /// Tracks window lifecycle events automatically
        /// </summary>
        public void AttachToWindow(Window window, string windowName = null)
        {
            if (window == null) return;
            
            var name = windowName ?? window.GetType().Name;
            
            window.Loaded += (s, e) =>
            {
                _windowOpenTimes[name] = DateTime.UtcNow;
                _telemetryService.TrackWindowEvent(name, "Opened", new Dictionary<string, object>
                {
                    ["WindowTitle"] = window.Title,
                    ["WindowType"] = window.GetType().Name,
                    ["Width"] = window.Width,
                    ["Height"] = window.Height
                });
            };
            
            window.Closed += (s, e) =>
            {
                if (_windowOpenTimes.TryGetValue(name, out var openTime))
                {
                    var duration = DateTime.UtcNow - openTime;
                    _telemetryService.TrackWindowEvent(name, "Closed", new Dictionary<string, object>
                    {
                        ["WindowTitle"] = window.Title,
                        ["WindowType"] = window.GetType().Name,
                        ["SessionDuration"] = duration.TotalMinutes
                    });
                    _windowOpenTimes.Remove(name);
                }
            };
            
            window.SizeChanged += (s, e) =>
            {
                _telemetryService.TrackWindowEvent(name, "Resized", new Dictionary<string, object>
                {
                    ["WindowTitle"] = window.Title,
                    ["NewWidth"] = e.NewSize.Width,
                    ["NewHeight"] = e.NewSize.Height,
                    ["OldWidth"] = e.PreviousSize.Width,
                    ["OldHeight"] = e.PreviousSize.Height
                });
            };
            
            window.StateChanged += (s, e) =>
            {
                _telemetryService.TrackWindowEvent(name, "StateChanged", new Dictionary<string, object>
                {
                    ["WindowTitle"] = window.Title,
                    ["NewState"] = window.WindowState.ToString(),
                    ["PreviousState"] = e.ToString()
                });
            };
        }

        /// <summary>
        /// Tracks button clicks with automatic context detection
        /// </summary>
        public void TrackButtonClick(string buttonName, string context = null, Dictionary<string, object> additionalData = null)
        {
            IncrementInteractionCount("Button_" + buttonName);
            
            var properties = new Dictionary<string, object>
            {
                ["ButtonName"] = buttonName,
                ["Context"] = context ?? "Unknown",
                ["InteractionCount"] = _uiInteractionCounts.GetValueOrDefault("Button_" + buttonName, 0)
            };
            
            if (additionalData != null)
            {
                foreach (var data in additionalData)
                {
                    properties[data.Key] = data.Value;
                }
            }
            
            _telemetryService.TrackUIInteraction("Button", buttonName, "Click", context);
        }

        /// <summary>
        /// Tracks navigation events between views or modes
        /// </summary>
        public void TrackNavigation(string fromView, string toView, TimeSpan? navigationTime = null)
        {
            var properties = new Dictionary<string, object>
            {
                ["FromView"] = fromView,
                ["ToView"] = toView,
                ["NavigationTime"] = navigationTime?.TotalMilliseconds ?? 0
            };
            
            _telemetryService.TrackEvent("Navigation", properties);
        }

        /// <summary>
        /// Tracks form submissions and validations
        /// </summary>
        public void TrackFormSubmission(string formName, bool isValid, int fieldCount, 
            Dictionary<string, object> validationErrors = null)
        {
            var properties = new Dictionary<string, object>
            {
                ["FormName"] = formName,
                ["IsValid"] = isValid,
                ["FieldCount"] = fieldCount,
                ["ValidationErrorCount"] = validationErrors?.Count ?? 0
            };
            
            if (validationErrors != null)
            {
                foreach (var error in validationErrors)
                {
                    properties[$"ValidationError_{error.Key}"] = error.Value;
                }
            }
            
            _telemetryService.TrackEvent("FormSubmission", properties);
        }

        /// <summary>
        /// Tracks drag and drop operations
        /// </summary>
        public void TrackDragDrop(string sourceType, string targetType, int itemCount, bool successful)
        {
            var properties = new Dictionary<string, object>
            {
                ["SourceType"] = sourceType,
                ["TargetType"] = targetType,
                ["ItemCount"] = itemCount,
                ["Successful"] = successful
            };
            
            _telemetryService.TrackEvent("DragDrop", properties);
        }

        /// <summary>
        /// Tracks application mode changes with timing
        /// </summary>
        public void TrackModeChange(string fromMode, string toMode, TimeSpan switchTime)
        {
            _telemetryService.TrackModeSwitch(fromMode, toMode, switchTime);
        }

        /// <summary>
        /// Tracks application startup with detailed context
        /// </summary>
        public void TrackApplicationStartup(int serviceCount, TimeSpan startupTime, string[] commandLineArgs = null)
        {
            var startupContext = new Dictionary<string, object>
            {
                ["CommandLineArgs"] = commandLineArgs != null ? string.Join(" ", commandLineArgs) : "",
                ["ArgumentCount"] = commandLineArgs?.Length ?? 0,
                ["MachineName"] = Environment.MachineName,
                ["UserName"] = Environment.UserName,
                ["WorkingDirectory"] = Environment.CurrentDirectory,
                ["ProcessorCount"] = Environment.ProcessorCount,
                ["OSVersion"] = Environment.OSVersion.ToString(),
                ["CLRVersion"] = Environment.Version.ToString()
            };
            
            _telemetryService.TrackApplicationStartup(startupTime, serviceCount, "WPF", startupContext);
        }

        private void IncrementInteractionCount(string key)
        {
            if (_uiInteractionCounts.ContainsKey(key))
            {
                _uiInteractionCounts[key]++;
            }
            else
            {
                _uiInteractionCounts[key] = 1;
            }
        }

        private void CollectPerformanceMetrics(object state)
        {
            try
            {
                using var process = Process.GetCurrentProcess();
                
                // Get CPU usage from performance counter (simplified approach)
                double cpuUsage = 0;
                try
                {
                    using var cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
                    cpuCounter.NextValue(); // First call returns 0, so we need to discard it
                    System.Threading.Thread.Sleep(100); // Wait a bit for accurate reading
                    cpuUsage = cpuCounter.NextValue();
                }
                catch
                {
                    cpuUsage = 0; // Fallback if performance counters fail
                }
                
                var memoryUsage = process.WorkingSet64;
                var threadCount = process.Threads.Count;
                var handleCount = process.HandleCount;
                
                _telemetryService.TrackPerformanceMetrics(cpuUsage, memoryUsage, threadCount, handleCount);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to collect performance metrics");
            }
        }

        private void CollectMemoryMetrics(object state)
        {
            try
            {
                // Force garbage collection to get accurate memory readings
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                
                var memoryBefore = GC.GetTotalMemory(false);
                var gen0Collections = GC.CollectionCount(0);
                var gen1Collections = GC.CollectionCount(1);
                var gen2Collections = GC.CollectionCount(2);
                
                var properties = new Dictionary<string, object>
                {
                    ["ManagedMemoryBytes"] = memoryBefore,
                    ["ManagedMemoryMB"] = memoryBefore / (1024.0 * 1024.0),
                    ["Gen0Collections"] = gen0Collections,
                    ["Gen1Collections"] = gen1Collections,
                    ["Gen2Collections"] = gen2Collections
                };
                
                var metrics = new Dictionary<string, double>
                {
                    ["ManagedMemoryBytes"] = memoryBefore,
                    ["Gen0Collections"] = gen0Collections,
                    ["Gen1Collections"] = gen1Collections,
                    ["Gen2Collections"] = gen2Collections
                };
                
                _telemetryService.TrackEvent("MemoryMetrics", properties, metrics);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to collect memory metrics");
            }
        }

        /// <summary>
        /// Gets session summary for telemetry reporting
        /// </summary>
        public Dictionary<string, object> GetSessionSummary()
        {
            var sessionDuration = DateTime.UtcNow - _sessionStart;
            var totalInteractions = 0;
            
            foreach (var count in _uiInteractionCounts.Values)
            {
                totalInteractions += count;
            }
            
            return new Dictionary<string, object>
            {
                ["SessionDuration"] = sessionDuration.TotalMinutes,
                ["TotalUIInteractions"] = totalInteractions,
                ["OpenWindowsCount"] = _windowOpenTimes.Count,
                ["UniqueInteractionTypes"] = _uiInteractionCounts.Count,
                ["SessionStart"] = _sessionStart,
                ["SessionEnd"] = DateTime.UtcNow
            };
        }

        public void Dispose()
        {
            if (_disposed) return;
            
            try
            {
                // Track session end
                var sessionSummary = GetSessionSummary();
                _telemetryService.TrackEvent("SessionEnd", sessionSummary);
                
                _performanceTimer?.Dispose();
                _memoryTimer?.Dispose();
                
                _logger.Information("WPF Telemetry Helper disposed");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during WPF Telemetry Helper disposal");
            }
            
            _disposed = true;
        }
    }
} 