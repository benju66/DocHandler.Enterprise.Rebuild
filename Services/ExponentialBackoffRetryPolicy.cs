using System;
using System.Threading.Tasks;
using DocHandler.Services.Configuration;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Implements exponential backoff retry policy
    /// </summary>
    public class ExponentialBackoffRetryPolicy : IRetryPolicy
    {
        private readonly ILogger _logger;
        private readonly SaveQuotesConfiguration _config;

        public ExponentialBackoffRetryPolicy(SaveQuotesConfiguration config)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            _logger = Log.ForContext<ExponentialBackoffRetryPolicy>();
        }

        public async Task<T> ExecuteAsync<T>(
            Func<Task<T>> operation,
            Action<Exception, int, TimeSpan>? onRetry = null)
        {
            if (operation == null) throw new ArgumentNullException(nameof(operation));

            var retryCount = 0;
            var delay = TimeSpan.FromMilliseconds(_config.RetryDelayMs);

            while (true)
            {
                try
                {
                    return await operation();
                }
                catch (Exception ex) when (ShouldRetry(ex) && retryCount < _config.MaxRetryAttempts)
                {
                    retryCount++;
                    
                    // Calculate next delay with exponential backoff
                    var nextDelay = TimeSpan.FromMilliseconds(
                        Math.Min(
                            delay.TotalMilliseconds * Math.Pow(_config.ExponentialBackoffFactor, retryCount - 1),
                            _config.MaxRetryDelayMs
                        )
                    );

                    _logger.Warning(ex, "Operation failed, retry {RetryCount}/{MaxRetries} after {Delay}ms",
                        retryCount, _config.MaxRetryAttempts, nextDelay.TotalMilliseconds);

                    // Notify caller about retry
                    onRetry?.Invoke(ex, retryCount, nextDelay);

                    // Wait before retry
                    await Task.Delay(nextDelay);
                }
                catch (Exception ex)
                {
                    // Max retries exceeded or non-retryable exception
                    if (retryCount >= _config.MaxRetryAttempts)
                    {
                        throw new MaxRetryExceededException(
                            $"Operation failed after {retryCount} retry attempts",
                            retryCount,
                            ex);
                    }
                    throw;
                }
            }
        }

        public async Task ExecuteAsync(
            Func<Task> operation,
            Action<Exception, int, TimeSpan>? onRetry = null)
        {
            await ExecuteAsync(async () =>
            {
                await operation();
                return true; // Dummy return value
            }, onRetry);
        }

        private bool ShouldRetry(Exception ex)
        {
            // Don't retry for circuit breaker open exceptions
            if (ex is CircuitBreakerOpenException)
                return false;

            // Don't retry for cancellation
            if (ex is OperationCanceledException)
                return false;

            // Don't retry for argument exceptions
            if (ex is ArgumentException)
                return false;

            // Retry for IO exceptions
            if (ex is System.IO.IOException)
                return true;

            // Retry for COM exceptions
            if (ex is System.Runtime.InteropServices.COMException)
                return true;

            // Retry for timeout exceptions
            if (ex is TimeoutException)
                return true;

            // Default: don't retry
            return false;
        }
    }
} 