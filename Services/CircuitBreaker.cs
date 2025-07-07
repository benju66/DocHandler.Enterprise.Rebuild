using System;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class CircuitBreaker
    {
        private static readonly ILogger _logger = Log.ForContext<CircuitBreaker>();
        
        private readonly object _lockObject = new object();
        private readonly int _failureThreshold;
        private readonly TimeSpan _recoveryTimeout;
        private readonly TimeSpan _successThreshold;
        
        private int _failureCount;
        private DateTime _lastFailureTime;
        private DateTime _lastSuccessTime;
        private CircuitBreakerState _state = CircuitBreakerState.Closed;
        
        public event EventHandler<CircuitBreakerStateChangedEventArgs> StateChanged;
        public event EventHandler<CircuitBreakerEventArgs> CircuitOpened;
        public event EventHandler<CircuitBreakerEventArgs> CircuitClosed;
        
        public CircuitBreaker(int failureThreshold = 5, TimeSpan? recoveryTimeout = null, TimeSpan? successThreshold = null)
        {
            _failureThreshold = failureThreshold;
            _recoveryTimeout = recoveryTimeout ?? TimeSpan.FromMinutes(1);
            _successThreshold = successThreshold ?? TimeSpan.FromMinutes(5);
            _lastSuccessTime = DateTime.UtcNow;
            
            _logger.Information("Circuit breaker initialized with threshold: {Threshold}, recovery timeout: {RecoveryTimeout}", 
                _failureThreshold, _recoveryTimeout);
        }
        
        public CircuitBreakerState State
        {
            get
            {
                lock (_lockObject)
                {
                    return _state;
                }
            }
        }
        
        public int FailureCount
        {
            get
            {
                lock (_lockObject)
                {
                    return _failureCount;
                }
            }
        }
        
        public DateTime LastFailureTime
        {
            get
            {
                lock (_lockObject)
                {
                    return _lastFailureTime;
                }
            }
        }
        
        /// <summary>
        /// Executes an operation through the circuit breaker
        /// </summary>
        public async Task<T> ExecuteAsync<T>(Func<Task<T>> operation, string operationName = null)
        {
            await CheckStateAsync();
            
            if (_state == CircuitBreakerState.Open)
            {
                var exception = new CircuitBreakerOpenException($"Circuit breaker is open. Operation '{operationName}' not executed.");
                _logger.Warning("Circuit breaker prevented execution of operation: {OperationName}", operationName);
                throw exception;
            }
            
            try
            {
                var result = await operation();
                await OnSuccessAsync(operationName);
                return result;
            }
            catch (Exception ex)
            {
                await OnFailureAsync(ex, operationName);
                throw;
            }
        }
        
        /// <summary>
        /// Executes a synchronous operation through the circuit breaker
        /// </summary>
        public T Execute<T>(Func<T> operation, string operationName = null)
        {
            CheckState();
            
            if (_state == CircuitBreakerState.Open)
            {
                var exception = new CircuitBreakerOpenException($"Circuit breaker is open. Operation '{operationName}' not executed.");
                _logger.Warning("Circuit breaker prevented execution of operation: {OperationName}", operationName);
                throw exception;
            }
            
            try
            {
                var result = operation();
                OnSuccess(operationName);
                return result;
            }
            catch (Exception ex)
            {
                OnFailure(ex, operationName);
                throw;
            }
        }
        
        /// <summary>
        /// Executes an action through the circuit breaker
        /// </summary>
        public void Execute(Action operation, string operationName = null)
        {
            CheckState();
            
            if (_state == CircuitBreakerState.Open)
            {
                var exception = new CircuitBreakerOpenException($"Circuit breaker is open. Operation '{operationName}' not executed.");
                _logger.Warning("Circuit breaker prevented execution of operation: {OperationName}", operationName);
                throw exception;
            }
            
            try
            {
                operation();
                OnSuccess(operationName);
            }
            catch (Exception ex)
            {
                OnFailure(ex, operationName);
                throw;
            }
        }
        
        /// <summary>
        /// Manually resets the circuit breaker
        /// </summary>
        public void Reset()
        {
            lock (_lockObject)
            {
                _failureCount = 0;
                _lastFailureTime = DateTime.MinValue;
                _lastSuccessTime = DateTime.UtcNow;
                ChangeState(CircuitBreakerState.Closed);
            }
            
            _logger.Information("Circuit breaker manually reset");
        }
        
        /// <summary>
        /// Gets current circuit breaker status
        /// </summary>
        public CircuitBreakerStatus GetStatus()
        {
            lock (_lockObject)
            {
                return new CircuitBreakerStatus
                {
                    State = _state,
                    FailureCount = _failureCount,
                    LastFailureTime = _lastFailureTime,
                    LastSuccessTime = _lastSuccessTime,
                    NextAttemptTime = _lastFailureTime.Add(_recoveryTimeout)
                };
            }
        }
        
        private async Task CheckStateAsync()
        {
            await Task.Run(() => CheckState());
        }
        
        private void CheckState()
        {
            lock (_lockObject)
            {
                if (_state == CircuitBreakerState.Open)
                {
                    if (DateTime.UtcNow >= _lastFailureTime.Add(_recoveryTimeout))
                    {
                        ChangeState(CircuitBreakerState.HalfOpen);
                        _logger.Information("Circuit breaker transitioning to half-open state");
                    }
                }
            }
        }
        
        private async Task OnSuccessAsync(string operationName)
        {
            await Task.Run(() => OnSuccess(operationName));
        }
        
        private void OnSuccess(string operationName)
        {
            lock (_lockObject)
            {
                var previousState = _state;
                _lastSuccessTime = DateTime.UtcNow;
                
                if (_state == CircuitBreakerState.HalfOpen)
                {
                    _failureCount = 0;
                    ChangeState(CircuitBreakerState.Closed);
                    _logger.Information("Circuit breaker closed after successful operation: {OperationName}", operationName);
                }
                else if (_state == CircuitBreakerState.Closed)
                {
                    // Reset failure count on successful operation
                    _failureCount = 0;
                }
                
                _logger.Debug("Circuit breaker operation succeeded: {OperationName}, State: {State}", operationName, _state);
            }
        }
        
        private async Task OnFailureAsync(Exception exception, string operationName)
        {
            await Task.Run(() => OnFailure(exception, operationName));
        }
        
        private void OnFailure(Exception exception, string operationName)
        {
            lock (_lockObject)
            {
                _failureCount++;
                _lastFailureTime = DateTime.UtcNow;
                
                _logger.Warning(exception, "Circuit breaker operation failed: {OperationName}, Failure count: {FailureCount}", 
                    operationName, _failureCount);
                
                if (_failureCount >= _failureThreshold)
                {
                    ChangeState(CircuitBreakerState.Open);
                    _logger.Error("Circuit breaker opened due to {FailureCount} failures. Operation: {OperationName}", 
                        _failureCount, operationName);
                }
            }
        }
        
        private void ChangeState(CircuitBreakerState newState)
        {
            var previousState = _state;
            _state = newState;
            
            StateChanged?.Invoke(this, new CircuitBreakerStateChangedEventArgs
            {
                PreviousState = previousState,
                CurrentState = newState,
                FailureCount = _failureCount,
                LastFailureTime = _lastFailureTime
            });
            
            if (newState == CircuitBreakerState.Open)
            {
                CircuitOpened?.Invoke(this, new CircuitBreakerEventArgs
                {
                    State = newState,
                    FailureCount = _failureCount,
                    LastFailureTime = _lastFailureTime
                });
            }
            else if (newState == CircuitBreakerState.Closed)
            {
                CircuitClosed?.Invoke(this, new CircuitBreakerEventArgs
                {
                    State = newState,
                    FailureCount = _failureCount,
                    LastFailureTime = _lastFailureTime
                });
            }
        }
    }
    
    public enum CircuitBreakerState
    {
        Closed,   // Normal operation
        Open,     // Failing fast
        HalfOpen  // Testing if service is back
    }
    
    public class CircuitBreakerStatus
    {
        public CircuitBreakerState State { get; set; }
        public int FailureCount { get; set; }
        public DateTime LastFailureTime { get; set; }
        public DateTime LastSuccessTime { get; set; }
        public DateTime NextAttemptTime { get; set; }
    }
    
    public class CircuitBreakerEventArgs : EventArgs
    {
        public CircuitBreakerState State { get; set; }
        public int FailureCount { get; set; }
        public DateTime LastFailureTime { get; set; }
    }
    
    public class CircuitBreakerStateChangedEventArgs : EventArgs
    {
        public CircuitBreakerState PreviousState { get; set; }
        public CircuitBreakerState CurrentState { get; set; }
        public int FailureCount { get; set; }
        public DateTime LastFailureTime { get; set; }
    }
    
    public class CircuitBreakerOpenException : Exception
    {
        public CircuitBreakerOpenException(string message) : base(message) { }
        public CircuitBreakerOpenException(string message, Exception innerException) : base(message, innerException) { }
    }
    
    /// <summary>
    /// Specialized circuit breaker for Office COM operations
    /// </summary>
    public class OfficeCircuitBreaker : CircuitBreaker
    {
        private static readonly ILogger _logger = Log.ForContext<OfficeCircuitBreaker>();
        
        public OfficeCircuitBreaker() : base(
            failureThreshold: 3,  // Lower threshold for Office operations
            recoveryTimeout: TimeSpan.FromMinutes(2),  // Longer recovery time for Office
            successThreshold: TimeSpan.FromMinutes(10))
        {
            _logger.Information("Office circuit breaker initialized with specialized settings");
        }
        
        /// <summary>
        /// Executes an Office COM operation with specialized error handling
        /// </summary>
        public async Task<T> ExecuteOfficeOperationAsync<T>(Func<Task<T>> operation, string operationName = null)
        {
            try
            {
                return await ExecuteAsync(operation, $"Office_{operationName}");
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                _logger.Error(comEx, "COM exception in Office operation: {OperationName}", operationName);
                
                // Consider COM exceptions as potential circuit breaker triggers
                if (IsRecoverableCOMException(comEx))
                {
                    _logger.Warning("Recoverable COM exception detected, will retry later: {OperationName}", operationName);
                    throw;
                }
                else
                {
                    _logger.Error("Non-recoverable COM exception, not triggering circuit breaker: {OperationName}", operationName);
                    throw;
                }
            }
        }
        
        /// <summary>
        /// Executes a synchronous Office COM operation
        /// </summary>
        public T ExecuteOfficeOperation<T>(Func<T> operation, string operationName = null)
        {
            try
            {
                return Execute(operation, $"Office_{operationName}");
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                _logger.Error(comEx, "COM exception in Office operation: {OperationName}", operationName);
                
                if (IsRecoverableCOMException(comEx))
                {
                    _logger.Warning("Recoverable COM exception detected, will retry later: {OperationName}", operationName);
                    throw;
                }
                else
                {
                    _logger.Error("Non-recoverable COM exception, not triggering circuit breaker: {OperationName}", operationName);
                    throw;
                }
            }
        }
        
        private bool IsRecoverableCOMException(System.Runtime.InteropServices.COMException comEx)
        {
            // Define which COM exceptions should trigger circuit breaker
            // These are typically transient errors that might be resolved by waiting
            var recoverableHResults = new[]
            {
                unchecked((int)0x800706BA), // RPC server unavailable
                unchecked((int)0x80010001), // Call was rejected by callee
                unchecked((int)0x80010105), // Server threw an exception
                unchecked((int)0x8001010A), // Message filter indicated application is busy
                unchecked((int)0x80004005), // Unspecified error (sometimes recoverable)
            };
            
            return Array.IndexOf(recoverableHResults, comEx.HResult) >= 0;
        }
    }
} 