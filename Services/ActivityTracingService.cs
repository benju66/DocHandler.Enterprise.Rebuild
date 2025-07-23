using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.DataContracts;
using Serilog;
using DocHandler.Services.Configuration;

namespace DocHandler.Services
{
    /// <summary>
    /// Provides activity tracing with correlation IDs for desktop workflows and user actions
    /// </summary>
    public class ActivityTracingService : IDisposable
    {
        private static readonly ILogger _logger = Log.ForContext<ActivityTracingService>();
        private static readonly ActivitySource _activitySource = new ActivitySource("DocHandler.Desktop");
        
        private readonly TelemetryService _telemetryService;
        private readonly IHierarchicalConfigurationService _configService;
        private readonly TelemetryClient _applicationInsightsClient;
        
        private readonly Dictionary<string, Activity> _activeActivities = new Dictionary<string, Activity>();
        private readonly Dictionary<string, DateTime> _activityStartTimes = new Dictionary<string, DateTime>();
        private readonly object _activitiesLock = new object();
        
        // Activity tracking statistics
        private int _totalActivities = 0;
        private int _completedActivities = 0;
        private int _failedActivities = 0;
        
        public ActivityTracingService(
            TelemetryService telemetryService,
            IHierarchicalConfigurationService configService = null,
            TelemetryClient applicationInsightsClient = null)
        {
            _telemetryService = telemetryService ?? throw new ArgumentNullException(nameof(telemetryService));
            _configService = configService;
            _applicationInsightsClient = applicationInsightsClient;
            
            _logger.Information("Activity Tracing Service initialized");
        }

        /// <summary>
        /// Starts a new workflow activity with correlation tracking
        /// </summary>
        public string StartWorkflowActivity(string workflowName, string workflowType, Dictionary<string, object> properties = null)
        {
            var activity = _activitySource.StartActivity($"Workflow_{workflowName}");
            if (activity == null)
            {
                _logger.Warning("Failed to start activity for workflow: {WorkflowName}", workflowName);
                return Guid.NewGuid().ToString(); // Fallback correlation ID
            }

            var correlationId = activity.Id ?? Guid.NewGuid().ToString();
            
            // Set activity tags
            activity.SetTag("workflow.name", workflowName);
            activity.SetTag("workflow.type", workflowType);
            activity.SetTag("correlation.id", correlationId);
            activity.SetTag("session.id", GetSessionId());
            activity.SetTag("user.id", Environment.UserName);
            activity.SetTag("machine.name", Environment.MachineName);
            
            if (properties != null)
            {
                foreach (var prop in properties)
                {
                    activity.SetTag($"custom.{prop.Key}", prop.Value?.ToString());
                }
            }

            lock (_activitiesLock)
            {
                _activeActivities[correlationId] = activity;
                _activityStartTimes[correlationId] = DateTime.UtcNow;
                Interlocked.Increment(ref _totalActivities);
            }

            // Track in telemetry
            _telemetryService.TrackEvent("WorkflowActivityStarted", new Dictionary<string, object>
            {
                ["WorkflowName"] = workflowName,
                ["WorkflowType"] = workflowType,
                ["CorrelationId"] = correlationId,
                ["ActivityId"] = activity.Id,
                ["ParentId"] = activity.ParentId,
                ["StartTime"] = DateTime.UtcNow
            });

            // Track in Application Insights if available
            if (_applicationInsightsClient != null && IsApplicationInsightsEnabled())
            {
                var requestTelemetry = new RequestTelemetry
                {
                    Name = $"{workflowType}_{workflowName}",
                    Id = correlationId,
                    Timestamp = DateTimeOffset.UtcNow,
                    Success = null // Will be set when completed
                };
                
                requestTelemetry.Properties["WorkflowName"] = workflowName;
                requestTelemetry.Properties["WorkflowType"] = workflowType;
                requestTelemetry.Properties["CorrelationId"] = correlationId;
                
                _applicationInsightsClient.TrackRequest(requestTelemetry);
            }

            _logger.Information("Started workflow activity: {WorkflowName} with correlation ID: {CorrelationId}", 
                workflowName, correlationId);

            return correlationId;
        }

        /// <summary>
        /// Adds a step or event to an existing workflow activity
        /// </summary>
        public void TrackActivityStep(string correlationId, string stepName, string stepType, 
            Dictionary<string, object> properties = null, bool isError = false)
        {
            Activity activity;
            lock (_activitiesLock)
            {
                if (!_activeActivities.TryGetValue(correlationId, out activity))
                {
                    _logger.Warning("Activity not found for correlation ID: {CorrelationId}", correlationId);
                    return;
                }
            }

            // Add event to activity
            var eventName = isError ? $"Error_{stepName}" : stepName;
            var activityEvent = new ActivityEvent(eventName, DateTimeOffset.UtcNow);
            
            if (properties != null)
            {
                var tags = new ActivityTagsCollection();
                foreach (var prop in properties)
                {
                    tags[prop.Key] = prop.Value?.ToString();
                }
                activityEvent = new ActivityEvent(eventName, DateTimeOffset.UtcNow, tags);
            }
            
            activity.AddEvent(activityEvent);

            // Track in telemetry
            var telemetryProperties = new Dictionary<string, object>
            {
                ["CorrelationId"] = correlationId,
                ["StepName"] = stepName,
                ["StepType"] = stepType,
                ["IsError"] = isError,
                ["ActivityId"] = activity.Id,
                ["Timestamp"] = DateTime.UtcNow
            };

            if (properties != null)
            {
                foreach (var prop in properties)
                {
                    telemetryProperties[$"Step_{prop.Key}"] = prop.Value;
                }
            }

            _telemetryService.TrackEvent("WorkflowActivityStep", telemetryProperties);

            // Track in Application Insights
            if (_applicationInsightsClient != null && IsApplicationInsightsEnabled())
            {
                var eventTelemetry = new EventTelemetry($"WorkflowStep_{stepName}")
                {
                    Timestamp = DateTimeOffset.UtcNow
                };
                
                eventTelemetry.Properties["CorrelationId"] = correlationId;
                eventTelemetry.Properties["StepName"] = stepName;
                eventTelemetry.Properties["StepType"] = stepType;
                eventTelemetry.Properties["IsError"] = isError.ToString();
                
                if (properties != null)
                {
                    foreach (var prop in properties)
                    {
                        eventTelemetry.Properties[prop.Key] = prop.Value?.ToString();
                    }
                }
                
                _applicationInsightsClient.TrackEvent(eventTelemetry);
            }

            _logger.Debug("Tracked activity step: {StepName} for correlation ID: {CorrelationId}", 
                stepName, correlationId);
        }

        /// <summary>
        /// Completes a workflow activity
        /// </summary>
        public void CompleteWorkflowActivity(string correlationId, bool successful, 
            Dictionary<string, object> completionData = null, string errorMessage = null)
        {
            Activity activity;
            DateTime startTime;
            
            lock (_activitiesLock)
            {
                if (!_activeActivities.TryGetValue(correlationId, out activity))
                {
                    _logger.Warning("Activity not found for completion: {CorrelationId}", correlationId);
                    return;
                }
                
                _activeActivities.Remove(correlationId);
                _activityStartTimes.TryGetValue(correlationId, out startTime);
                _activityStartTimes.Remove(correlationId);
                
                if (successful)
                {
                    Interlocked.Increment(ref _completedActivities);
                }
                else
                {
                    Interlocked.Increment(ref _failedActivities);
                }
            }

            var duration = DateTime.UtcNow - startTime;
            
            // Set activity result
            activity.SetStatus(successful ? ActivityStatusCode.Ok : ActivityStatusCode.Error, errorMessage);
            activity.SetTag("success", successful.ToString());
            activity.SetTag("duration.ms", duration.TotalMilliseconds.ToString());
            
            if (!string.IsNullOrEmpty(errorMessage))
            {
                activity.SetTag("error.message", errorMessage);
            }

            // Track completion in telemetry
            var telemetryProperties = new Dictionary<string, object>
            {
                ["CorrelationId"] = correlationId,
                ["Successful"] = successful,
                ["Duration"] = duration.TotalMilliseconds,
                ["ActivityId"] = activity.Id,
                ["CompletionTime"] = DateTime.UtcNow
            };

            if (!string.IsNullOrEmpty(errorMessage))
            {
                telemetryProperties["ErrorMessage"] = errorMessage;
            }

            if (completionData != null)
            {
                foreach (var data in completionData)
                {
                    telemetryProperties[$"Completion_{data.Key}"] = data.Value;
                }
            }

            _telemetryService.TrackEvent("WorkflowActivityCompleted", telemetryProperties);

            // Track in Application Insights
            if (_applicationInsightsClient != null && IsApplicationInsightsEnabled())
            {
                var requestTelemetry = new RequestTelemetry
                {
                    Name = activity.DisplayName,
                    Id = correlationId,
                    Duration = duration,
                    Success = successful,
                    Timestamp = DateTimeOffset.UtcNow.Subtract(duration)
                };
                
                requestTelemetry.Properties["CorrelationId"] = correlationId;
                requestTelemetry.Properties["Successful"] = successful.ToString();
                
                if (!string.IsNullOrEmpty(errorMessage))
                {
                    requestTelemetry.Properties["ErrorMessage"] = errorMessage;
                }
                
                if (completionData != null)
                {
                    foreach (var data in completionData)
                    {
                        requestTelemetry.Properties[data.Key] = data.Value?.ToString();
                    }
                }
                
                _applicationInsightsClient.TrackRequest(requestTelemetry);
            }

            // Dispose activity
            activity.Dispose();

            _logger.Information("Completed workflow activity with correlation ID: {CorrelationId}, Success: {Successful}, Duration: {Duration}ms", 
                correlationId, successful, duration.TotalMilliseconds);
        }

        /// <summary>
        /// Creates a correlation ID for linking related activities
        /// </summary>
        public string CreateCorrelationId()
        {
            return Guid.NewGuid().ToString();
        }

        /// <summary>
        /// Gets the current activity correlation ID if available
        /// </summary>
        public string GetCurrentCorrelationId()
        {
            var currentActivity = Activity.Current;
            if (currentActivity != null)
            {
                return currentActivity.Id ?? currentActivity.TraceId.ToString();
            }
            
            return null;
        }

        /// <summary>
        /// Tracks a dependency call (external service, database, file system)
        /// </summary>
        public void TrackDependency(string correlationId, string dependencyName, string dependencyType, 
            string target, TimeSpan duration, bool successful, string data = null)
        {
            var telemetryProperties = new Dictionary<string, object>
            {
                ["CorrelationId"] = correlationId,
                ["DependencyName"] = dependencyName,
                ["DependencyType"] = dependencyType,
                ["Target"] = target,
                ["Duration"] = duration.TotalMilliseconds,
                ["Successful"] = successful
            };

            if (!string.IsNullOrEmpty(data))
            {
                telemetryProperties["Data"] = data;
            }

            _telemetryService.TrackEvent("DependencyCall", telemetryProperties);

            // Track in Application Insights
            if (_applicationInsightsClient != null && IsApplicationInsightsEnabled())
            {
                var dependencyTelemetry = new DependencyTelemetry
                {
                    Name = dependencyName,
                    Type = dependencyType,
                    Target = target,
                    Duration = duration,
                    Success = successful,
                    Data = data,
                    Timestamp = DateTimeOffset.UtcNow.Subtract(duration)
                };
                
                dependencyTelemetry.Properties["CorrelationId"] = correlationId;
                
                _applicationInsightsClient.TrackDependency(dependencyTelemetry);
            }

            _logger.Debug("Tracked dependency: {DependencyName} ({DependencyType}) for correlation ID: {CorrelationId}", 
                dependencyName, dependencyType, correlationId);
        }

        /// <summary>
        /// Gets activity tracing statistics
        /// </summary>
        public Dictionary<string, object> GetActivityStatistics()
        {
            lock (_activitiesLock)
            {
                return new Dictionary<string, object>
                {
                    ["TotalActivities"] = _totalActivities,
                    ["CompletedActivities"] = _completedActivities,
                    ["FailedActivities"] = _failedActivities,
                    ["ActiveActivities"] = _activeActivities.Count,
                    ["SuccessRate"] = _totalActivities > 0 ? (double)_completedActivities / _totalActivities * 100 : 0,
                    ["FailureRate"] = _totalActivities > 0 ? (double)_failedActivities / _totalActivities * 100 : 0
                };
            }
        }

        private bool IsApplicationInsightsEnabled()
        {
            return _configService?.Config?.Telemetry?.EnableApplicationInsights == true;
        }

        private string GetSessionId()
        {
            // Try to get session ID from current activity or generate one
            var currentActivity = Activity.Current;
            if (currentActivity != null)
            {
                var sessionTag = currentActivity.GetTagItem("session.id");
                if (sessionTag != null)
                {
                    return sessionTag.ToString();
                }
            }
            
            // Fallback to process ID + start time
            var process = Process.GetCurrentProcess();
            return $"{process.Id}_{process.StartTime:yyyyMMddHHmmss}";
        }

        public void Dispose()
        {
            try
            {
                // Complete any remaining activities
                lock (_activitiesLock)
                {
                    foreach (var kvp in _activeActivities)
                    {
                        kvp.Value.SetStatus(ActivityStatusCode.Error, "Application shutdown");
                        kvp.Value.Dispose();
                    }
                    _activeActivities.Clear();
                    _activityStartTimes.Clear();
                }

                // Track final statistics
                var finalStats = GetActivityStatistics();
                _telemetryService.TrackEvent("ActivityTracingServiceDisposed", finalStats);

                _activitySource.Dispose();
                
                _logger.Information("Activity Tracing Service disposed. Final statistics: {@Statistics}", finalStats);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during Activity Tracing Service disposal");
            }
        }
    }

    /// <summary>
    /// Helper class for scoped activity tracking with using statement
    /// </summary>
    public class ScopedActivity : IDisposable
    {
        private readonly ActivityTracingService _tracingService;
        private readonly string _correlationId;
        private bool _disposed = false;

        public string CorrelationId => _correlationId;

        public ScopedActivity(ActivityTracingService tracingService, string workflowName, string workflowType, 
            Dictionary<string, object> properties = null)
        {
            _tracingService = tracingService;
            _correlationId = tracingService.StartWorkflowActivity(workflowName, workflowType, properties);
        }

        public void TrackStep(string stepName, string stepType, Dictionary<string, object> properties = null, bool isError = false)
        {
            _tracingService.TrackActivityStep(_correlationId, stepName, stepType, properties, isError);
        }

        public void TrackDependency(string dependencyName, string dependencyType, string target, 
            TimeSpan duration, bool successful, string data = null)
        {
            _tracingService.TrackDependency(_correlationId, dependencyName, dependencyType, target, duration, successful, data);
        }

        public void Complete(bool successful, Dictionary<string, object> completionData = null, string errorMessage = null)
        {
            if (!_disposed)
            {
                _tracingService.CompleteWorkflowActivity(_correlationId, successful, completionData, errorMessage);
                _disposed = true;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                Complete(true); // Default to successful completion
            }
        }
    }
} 