using System;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Circuit breaker pattern to prevent cascading failures in Word conversions
    /// </summary>
    public class ConversionCircuitBreaker
    {
        private readonly ILogger _logger = Log.ForContext<ConversionCircuitBreaker>();
        private readonly int _failureThreshold;
        private readonly TimeSpan _breakDuration;
        private readonly object _lock = new object();
        
        private int _failureCount = 0;
        private DateTime _lastFailureTime = DateTime.MinValue;
        private bool _isOpen = false;
        
        public ConversionCircuitBreaker(int failureThreshold = 5, TimeSpan? breakDuration = null)
        {
            _failureThreshold = failureThreshold;
            _breakDuration = breakDuration ?? TimeSpan.FromMinutes(2);
            
            _logger.Information("Circuit breaker initialized with {Threshold} failure threshold and {Duration} break duration",
                _failureThreshold, _breakDuration.TotalMinutes);
        }
        
        public bool IsOpen
        {
            get
            {
                lock (_lock)
                {
                    // Check if enough time has passed to close the circuit
                    if (_isOpen && DateTime.UtcNow - _lastFailureTime > _breakDuration)
                    {
                        _logger.Information("Circuit breaker auto-closing after {Duration} minutes", _breakDuration.TotalMinutes);
                        Reset();
                    }
                    
                    return _isOpen;
                }
            }
        }
        
        public async Task<T> ExecuteAsync<T>(Func<Task<T>> operation) where T : class
        {
            if (IsOpen)
            {
                _logger.Warning("Circuit breaker is OPEN - rejecting operation");
                throw new InvalidOperationException("Circuit breaker is open - too many recent failures");
            }
            
            try
            {
                var result = await operation();
                OnSuccess();
                return result;
            }
            catch (Exception ex)
            {
                OnFailure(ex);
                throw;
            }
        }
        
        private void OnSuccess()
        {
            lock (_lock)
            {
                if (_failureCount > 0)
                {
                    _logger.Debug("Operation succeeded - resetting failure count from {Count}", _failureCount);
                    _failureCount = 0;
                }
            }
        }
        
        private void OnFailure(Exception ex)
        {
            lock (_lock)
            {
                _failureCount++;
                _lastFailureTime = DateTime.UtcNow;
                
                _logger.Warning(ex, "Operation failed - failure count: {Count}/{Threshold}", 
                    _failureCount, _failureThreshold);
                
                if (_failureCount >= _failureThreshold)
                {
                    _isOpen = true;
                    _logger.Error("Circuit breaker OPENED after {Count} failures - will retry in {Duration} minutes",
                        _failureCount, _breakDuration.TotalMinutes);
                }
            }
        }
        
        public void Reset()
        {
            lock (_lock)
            {
                _failureCount = 0;
                _isOpen = false;
                _lastFailureTime = DateTime.MinValue;
                _logger.Information("Circuit breaker manually reset");
            }
        }
        
        public (int failures, bool isOpen, TimeSpan timeUntilClose) GetStatus()
        {
            lock (_lock)
            {
                var timeUntilClose = _isOpen 
                    ? _breakDuration - (DateTime.UtcNow - _lastFailureTime)
                    : TimeSpan.Zero;
                    
                return (_failureCount, _isOpen, timeUntilClose);
            }
        }
    }
} 