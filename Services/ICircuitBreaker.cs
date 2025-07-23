using System;
using System.Threading.Tasks;

namespace DocHandler.Services
{
    /// <summary>
    /// Defines a circuit breaker for fault tolerance
    /// </summary>
    public interface ICircuitBreaker
    {
        /// <summary>
        /// Current state of the circuit breaker
        /// </summary>
        CircuitBreakerState State { get; }

        /// <summary>
        /// Execute an operation through the circuit breaker
        /// </summary>
        Task<T> ExecuteAsync<T>(Func<Task<T>> operation);

        /// <summary>
        /// Execute an operation through the circuit breaker (void return)
        /// </summary>
        Task ExecuteAsync(Func<Task> operation);

        /// <summary>
        /// Manually reset the circuit breaker
        /// </summary>
        void Reset();

        /// <summary>
        /// Event raised when circuit breaker state changes
        /// </summary>
        event EventHandler<CircuitBreakerStateChangedEventArgs> StateChanged;
    }

    /// <summary>
    /// Circuit breaker states
    /// </summary>
    public enum CircuitBreakerState
    {
        Closed,
        Open,
        HalfOpen
    }

    /// <summary>
    /// Event args for circuit breaker state changes
    /// </summary>
    public class CircuitBreakerStateChangedEventArgs : EventArgs
    {
        public CircuitBreakerState OldState { get; }
        public CircuitBreakerState NewState { get; }
        public string Reason { get; }

        public CircuitBreakerStateChangedEventArgs(CircuitBreakerState oldState, CircuitBreakerState newState, string reason)
        {
            OldState = oldState;
            NewState = newState;
            Reason = reason;
        }
    }
} 