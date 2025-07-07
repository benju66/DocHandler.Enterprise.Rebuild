using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class RetryPolicies
    {
        private static readonly ILogger _logger = Log.ForContext<RetryPolicies>();

        /// <summary>
        /// Executes an operation with retry logic using exponential backoff
        /// </summary>
        public static async Task<T> ExecuteWithRetryAsync<T>(
            Func<Task<T>> operation,
            int maxRetries = 3,
            TimeSpan? initialDelay = null,
            double backoffMultiplier = 2.0,
            TimeSpan? maxDelay = null,
            Func<Exception, bool> shouldRetry = null,
            string operationName = null)
        {
            var delay = initialDelay ?? TimeSpan.FromSeconds(1);
            var maxDelayValue = maxDelay ?? TimeSpan.FromMinutes(1);
            var attempt = 0;
            
            while (true)
            {
                try
                {
                    var result = await operation();
                    
                    if (attempt > 0)
                    {
                        _logger.Information("Operation succeeded after {Attempts} attempts: {OperationName}", 
                            attempt + 1, operationName);
                    }
                    
                    return result;
                }
                catch (Exception ex)
                {
                    attempt++;
                    
                    // Check if we should retry this exception
                    var shouldRetryException = shouldRetry?.Invoke(ex) ?? IsRetryableException(ex);
                    
                    if (attempt > maxRetries || !shouldRetryException)
                    {
                        _logger.Error(ex, "Operation failed after {Attempts} attempts: {OperationName}", 
                            attempt, operationName);
                        throw;
                    }
                    
                    // Calculate delay with exponential backoff
                    var currentDelay = TimeSpan.FromMilliseconds(
                        Math.Min(delay.TotalMilliseconds * Math.Pow(backoffMultiplier, attempt - 1), 
                                maxDelayValue.TotalMilliseconds));
                    
                    _logger.Warning(ex, "Operation failed (attempt {Attempt}/{MaxRetries}), retrying in {Delay}ms: {OperationName}",
                        attempt, maxRetries, currentDelay.TotalMilliseconds, operationName);
                    
                    await Task.Delay(currentDelay);
                }
            }
        }
        
        /// <summary>
        /// Executes a synchronous operation with retry logic
        /// </summary>
        public static T ExecuteWithRetry<T>(
            Func<T> operation,
            int maxRetries = 3,
            TimeSpan? initialDelay = null,
            double backoffMultiplier = 2.0,
            TimeSpan? maxDelay = null,
            Func<Exception, bool> shouldRetry = null,
            string operationName = null)
        {
            var delay = initialDelay ?? TimeSpan.FromSeconds(1);
            var maxDelayValue = maxDelay ?? TimeSpan.FromMinutes(1);
            var attempt = 0;
            
            while (true)
            {
                try
                {
                    var result = operation();
                    
                    if (attempt > 0)
                    {
                        _logger.Information("Operation succeeded after {Attempts} attempts: {OperationName}", 
                            attempt + 1, operationName);
                    }
                    
                    return result;
                }
                catch (Exception ex)
                {
                    attempt++;
                    
                    // Check if we should retry this exception
                    var shouldRetryException = shouldRetry?.Invoke(ex) ?? IsRetryableException(ex);
                    
                    if (attempt > maxRetries || !shouldRetryException)
                    {
                        _logger.Error(ex, "Operation failed after {Attempts} attempts: {OperationName}", 
                            attempt, operationName);
                        throw;
                    }
                    
                    // Calculate delay with exponential backoff
                    var currentDelay = TimeSpan.FromMilliseconds(
                        Math.Min(delay.TotalMilliseconds * Math.Pow(backoffMultiplier, attempt - 1), 
                                maxDelayValue.TotalMilliseconds));
                    
                    _logger.Warning(ex, "Operation failed (attempt {Attempt}/{MaxRetries}), retrying in {Delay}ms: {OperationName}",
                        attempt, maxRetries, currentDelay.TotalMilliseconds, operationName);
                    
                    Thread.Sleep(currentDelay);
                }
            }
        }
        
        /// <summary>
        /// Executes an action with retry logic
        /// </summary>
        public static void ExecuteWithRetry(
            Action operation,
            int maxRetries = 3,
            TimeSpan? initialDelay = null,
            double backoffMultiplier = 2.0,
            TimeSpan? maxDelay = null,
            Func<Exception, bool> shouldRetry = null,
            string operationName = null)
        {
            ExecuteWithRetry<object>(() =>
            {
                operation();
                return null;
            }, maxRetries, initialDelay, backoffMultiplier, maxDelay, shouldRetry, operationName);
        }
        
        /// <summary>
        /// Determines if an exception is retryable
        /// </summary>
        public static bool IsRetryableException(Exception ex)
        {
            // Common retryable exceptions
            var retryableExceptions = new[]
            {
                typeof(TimeoutException),
                typeof(TaskCanceledException),
                typeof(OperationCanceledException),
                typeof(System.Net.WebException),
                typeof(System.Net.Http.HttpRequestException),
                typeof(System.IO.IOException),
                typeof(System.IO.FileNotFoundException),
                typeof(System.IO.DirectoryNotFoundException),
                typeof(UnauthorizedAccessException),
                typeof(System.Runtime.InteropServices.COMException)
            };
            
            var exceptionType = ex.GetType();
            
            // Check direct type match
            if (retryableExceptions.Contains(exceptionType))
            {
                return true;
            }
            
            // Check for specific COM exception error codes
            if (ex is System.Runtime.InteropServices.COMException comEx)
            {
                return IsRetryableCOMException(comEx);
            }
            
            // Check for specific IO exceptions
            if (ex is System.IO.IOException ioEx)
            {
                return IsRetryableIOException(ioEx);
            }
            
            // Check inner exceptions
            if (ex.InnerException != null)
            {
                return IsRetryableException(ex.InnerException);
            }
            
            return false;
        }
        
        private static bool IsRetryableCOMException(System.Runtime.InteropServices.COMException comEx)
        {
            // Retryable COM exception HRESULTs
            var retryableHResults = new[]
            {
                unchecked((int)0x800706BA), // RPC server unavailable
                unchecked((int)0x80010001), // Call was rejected by callee
                unchecked((int)0x80010105), // Server threw an exception
                unchecked((int)0x8001010A), // Message filter indicated application is busy
                unchecked((int)0x80004005), // Unspecified error (sometimes recoverable)
                unchecked((int)0x80070006), // Invalid handle
                unchecked((int)0x800706BE), // Remote procedure call failed
            };
            
            return retryableHResults.Contains(comEx.HResult);
        }
        
        private static bool IsRetryableIOException(System.IO.IOException ioEx)
        {
            // File in use, sharing violation, etc.
            var message = ioEx.Message?.ToLowerInvariant() ?? "";
            
            var retryableMessages = new[]
            {
                "being used by another process",
                "sharing violation",
                "file is locked",
                "access denied",
                "device not ready"
            };
            
            return retryableMessages.Any(msg => message.Contains(msg));
        }
    }
    
    /// <summary>
    /// Predefined retry policies for common scenarios
    /// </summary>
    public static class RetryPolicyPresets
    {
        /// <summary>
        /// Aggressive retry policy for critical operations
        /// </summary>
        public static RetryPolicy Aggressive => new RetryPolicy
        {
            MaxRetries = 5,
            InitialDelay = TimeSpan.FromMilliseconds(500),
            BackoffMultiplier = 2.0,
            MaxDelay = TimeSpan.FromMinutes(2),
            ShouldRetry = RetryPolicies.IsRetryableException
        };
        
        /// <summary>
        /// Conservative retry policy for non-critical operations
        /// </summary>
        public static RetryPolicy Conservative => new RetryPolicy
        {
            MaxRetries = 2,
            InitialDelay = TimeSpan.FromSeconds(1),
            BackoffMultiplier = 1.5,
            MaxDelay = TimeSpan.FromSeconds(30),
            ShouldRetry = RetryPolicies.IsRetryableException
        };
        
        /// <summary>
        /// Office-specific retry policy for COM operations
        /// </summary>
        public static RetryPolicy Office => new RetryPolicy
        {
            MaxRetries = 3,
            InitialDelay = TimeSpan.FromSeconds(2),
            BackoffMultiplier = 2.0,
            MaxDelay = TimeSpan.FromMinutes(1),
            ShouldRetry = ex => ex is System.Runtime.InteropServices.COMException || 
                               RetryPolicies.IsRetryableException(ex)
        };
        
        /// <summary>
        /// File I/O specific retry policy
        /// </summary>
        public static RetryPolicy FileIO => new RetryPolicy
        {
            MaxRetries = 4,
            InitialDelay = TimeSpan.FromMilliseconds(250),
            BackoffMultiplier = 1.5,
            MaxDelay = TimeSpan.FromSeconds(15),
            ShouldRetry = ex => ex is System.IO.IOException || 
                               ex is UnauthorizedAccessException ||
                               RetryPolicies.IsRetryableException(ex)
        };
        
        /// <summary>
        /// Network operation retry policy
        /// </summary>
        public static RetryPolicy Network => new RetryPolicy
        {
            MaxRetries = 3,
            InitialDelay = TimeSpan.FromSeconds(1),
            BackoffMultiplier = 2.0,
            MaxDelay = TimeSpan.FromSeconds(30),
            ShouldRetry = ex => ex is System.Net.WebException || 
                               ex is System.Net.Http.HttpRequestException ||
                               ex is TimeoutException ||
                               RetryPolicies.IsRetryableException(ex)
        };
    }
    
    /// <summary>
    /// Retry policy configuration
    /// </summary>
    public class RetryPolicy
    {
        public int MaxRetries { get; set; } = 3;
        public TimeSpan InitialDelay { get; set; } = TimeSpan.FromSeconds(1);
        public double BackoffMultiplier { get; set; } = 2.0;
        public TimeSpan MaxDelay { get; set; } = TimeSpan.FromMinutes(1);
        public Func<Exception, bool> ShouldRetry { get; set; } = RetryPolicies.IsRetryableException;
        
        /// <summary>
        /// Executes an async operation with this retry policy
        /// </summary>
        public async Task<T> ExecuteAsync<T>(Func<Task<T>> operation, string operationName = null)
        {
            return await RetryPolicies.ExecuteWithRetryAsync(
                operation, MaxRetries, InitialDelay, BackoffMultiplier, MaxDelay, ShouldRetry, operationName);
        }
        
        /// <summary>
        /// Executes a synchronous operation with this retry policy
        /// </summary>
        public T Execute<T>(Func<T> operation, string operationName = null)
        {
            return RetryPolicies.ExecuteWithRetry(
                operation, MaxRetries, InitialDelay, BackoffMultiplier, MaxDelay, ShouldRetry, operationName);
        }
        
        /// <summary>
        /// Executes an action with this retry policy
        /// </summary>
        public void Execute(Action operation, string operationName = null)
        {
            RetryPolicies.ExecuteWithRetry(
                operation, MaxRetries, InitialDelay, BackoffMultiplier, MaxDelay, ShouldRetry, operationName);
        }
    }
    
    /// <summary>
    /// Retry policy builder for creating custom policies
    /// </summary>
    public class RetryPolicyBuilder
    {
        private readonly RetryPolicy _policy = new RetryPolicy();
        
        public RetryPolicyBuilder WithMaxRetries(int maxRetries)
        {
            _policy.MaxRetries = maxRetries;
            return this;
        }
        
        public RetryPolicyBuilder WithInitialDelay(TimeSpan delay)
        {
            _policy.InitialDelay = delay;
            return this;
        }
        
        public RetryPolicyBuilder WithBackoffMultiplier(double multiplier)
        {
            _policy.BackoffMultiplier = multiplier;
            return this;
        }
        
        public RetryPolicyBuilder WithMaxDelay(TimeSpan maxDelay)
        {
            _policy.MaxDelay = maxDelay;
            return this;
        }
        
        public RetryPolicyBuilder WithCustomRetryCondition(Func<Exception, bool> shouldRetry)
        {
            _policy.ShouldRetry = shouldRetry;
            return this;
        }
        
        public RetryPolicyBuilder ForFileOperations()
        {
            _policy.MaxRetries = 4;
            _policy.InitialDelay = TimeSpan.FromMilliseconds(250);
            _policy.BackoffMultiplier = 1.5;
            _policy.MaxDelay = TimeSpan.FromSeconds(15);
            _policy.ShouldRetry = ex => ex is System.IO.IOException || 
                                      ex is UnauthorizedAccessException ||
                                      RetryPolicies.IsRetryableException(ex);
            return this;
        }
        
        public RetryPolicyBuilder ForOfficeOperations()
        {
            _policy.MaxRetries = 3;
            _policy.InitialDelay = TimeSpan.FromSeconds(2);
            _policy.BackoffMultiplier = 2.0;
            _policy.MaxDelay = TimeSpan.FromMinutes(1);
            _policy.ShouldRetry = ex => ex is System.Runtime.InteropServices.COMException ||
                                      RetryPolicies.IsRetryableException(ex);
            return this;
        }
        
        public RetryPolicyBuilder ForNetworkOperations()
        {
            _policy.MaxRetries = 3;
            _policy.InitialDelay = TimeSpan.FromSeconds(1);
            _policy.BackoffMultiplier = 2.0;
            _policy.MaxDelay = TimeSpan.FromSeconds(30);
            _policy.ShouldRetry = ex => ex is System.Net.WebException ||
                                      ex is System.Net.Http.HttpRequestException ||
                                      ex is TimeoutException ||
                                      RetryPolicies.IsRetryableException(ex);
            return this;
        }
        
        public RetryPolicy Build()
        {
            return _policy;
        }
    }
} 