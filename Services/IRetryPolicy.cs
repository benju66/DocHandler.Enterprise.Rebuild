using System;
using System.Threading.Tasks;

namespace DocHandler.Services
{
    /// <summary>
    /// Defines a retry policy for operations that may fail transiently
    /// </summary>
    public interface IRetryPolicy
    {
        /// <summary>
        /// Execute an operation with retry logic
        /// </summary>
        Task<T> ExecuteAsync<T>(
            Func<Task<T>> operation,
            Action<Exception, int, TimeSpan>? onRetry = null);

        /// <summary>
        /// Execute an operation with retry logic (void return)
        /// </summary>
        Task ExecuteAsync(
            Func<Task> operation,
            Action<Exception, int, TimeSpan>? onRetry = null);
    }

    /// <summary>
    /// Exception thrown when retry attempts are exhausted
    /// </summary>
    public class MaxRetryExceededException : Exception
    {
        public int RetryCount { get; }

        public MaxRetryExceededException(string message, int retryCount, Exception innerException)
            : base(message, innerException)
        {
            RetryCount = retryCount;
        }
    }
    
    /// <summary>
    /// Exception thrown when circuit breaker is open
    /// </summary>
    public class CircuitBreakerOpenException : Exception
    {
        public TimeSpan RetryAfter { get; }

        public CircuitBreakerOpenException(string message, TimeSpan retryAfter)
            : base(message)
        {
            RetryAfter = retryAfter;
        }
    }
} 