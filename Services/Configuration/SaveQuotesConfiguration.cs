using System;
using System.ComponentModel;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Configuration settings specific to SaveQuotes processing mode
    /// </summary>
    public class SaveQuotesConfiguration
    {
        /// <summary>
        /// The default processing mode for SaveQuotes
        /// </summary>
        [DefaultValue(ProcessingMode.Pipeline)]
        public ProcessingMode DefaultProcessingMode { get; set; } = ProcessingMode.Pipeline;

        /// <summary>
        /// Enable security validation for all input files
        /// </summary>
        [DefaultValue(true)]
        public bool EnableSecurityValidation { get; set; } = true;

        /// <summary>
        /// Maximum number of retry attempts for failed conversions
        /// </summary>
        [DefaultValue(3)]
        public int MaxRetryAttempts { get; set; } = 3;

        /// <summary>
        /// Initial delay in milliseconds between retry attempts
        /// </summary>
        [DefaultValue(1000)]
        public int RetryDelayMs { get; set; } = 1000;

        /// <summary>
        /// Maximum delay in milliseconds between retry attempts
        /// </summary>
        [DefaultValue(30000)]
        public int MaxRetryDelayMs { get; set; } = 30000;

        /// <summary>
        /// Exponential backoff factor for retry delays
        /// </summary>
        [DefaultValue(2.0)]
        public double ExponentialBackoffFactor { get; set; } = 2.0;

        /// <summary>
        /// Enable batch processing for multiple files
        /// </summary>
        [DefaultValue(true)]
        public bool EnableBatchProcessing { get; set; } = true;

        /// <summary>
        /// Number of files to process in each batch
        /// </summary>
        [DefaultValue(10)]
        public int BatchSize { get; set; } = 10;

        /// <summary>
        /// Maximum concurrent file processing operations
        /// </summary>
        [DefaultValue(4)]
        public int MaxConcurrency { get; set; } = 4;

        /// <summary>
        /// Enable verbose logging for debugging
        /// </summary>
        [DefaultValue(false)]
        public bool EnableVerboseLogging { get; set; } = false;
        
        /// <summary>
        /// Log every Nth file to reduce log volume (1 = log every file, 10 = log every 10th file)
        /// </summary>
        [DefaultValue(10)]
        public int LogEveryNthFile { get; set; } = 10;

        /// <summary>
        /// Threshold in milliseconds for logging slow operations
        /// </summary>
        [DefaultValue(1000)]
        public int SlowOperationThresholdMs { get; set; } = 1000;

        /// <summary>
        /// Maximum cache size in MB for PDF conversions
        /// </summary>
        [DefaultValue(500)]
        public long MaxCacheSizeMB { get; set; } = 500;

        /// <summary>
        /// Cache sliding expiration in minutes
        /// </summary>
        [DefaultValue(30)]
        public int CacheSlidingExpirationMinutes { get; set; } = 30;

        /// <summary>
        /// Enable circuit breaker for conversion operations
        /// </summary>
        [DefaultValue(true)]
        public bool EnableCircuitBreaker { get; set; } = true;

        /// <summary>
        /// Circuit breaker failure threshold
        /// </summary>
        [DefaultValue(5)]
        public int CircuitBreakerFailureThreshold { get; set; } = 5;

        /// <summary>
        /// Circuit breaker sampling duration
        /// </summary>
        [DefaultValue("00:01:00")]
        public TimeSpan CircuitBreakerSamplingDuration { get; set; } = TimeSpan.FromMinutes(1);

        /// <summary>
        /// Circuit breaker break duration
        /// </summary>
        [DefaultValue("00:00:30")]
        public TimeSpan CircuitBreakerBreakDuration { get; set; } = TimeSpan.FromSeconds(30);

        /// <summary>
        /// Enable fallback to queue processing if pipeline fails
        /// </summary>
        [DefaultValue(false)]
        public bool EnableQueueFallback { get; set; } = false;
    }

    /// <summary>
    /// Processing mode for SaveQuotes
    /// </summary>
    public enum ProcessingMode
    {
        /// <summary>
        /// Use the secure pipeline architecture with validation
        /// </summary>
        Pipeline,

        /// <summary>
        /// Use legacy queue-based processing
        /// </summary>
        Queue,

        /// <summary>
        /// Use pipeline with queue for background processing
        /// </summary>
        Hybrid
    }
} 