using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Channels;
using System.Threading.Tasks;
using Serilog;
using System.Linq; // Added for .Count()

namespace DocHandler.Services
{
    /// <summary>
    /// A thread pool specifically designed for COM operations that require STA (Single Threaded Apartment) threading.
    /// This ensures that all COM object creation and access happens on dedicated STA threads, preventing
    /// RPC_E_CANTCALLOUT_ININPUTSYNCCALL errors and UI freezing.
    /// </summary>
    public sealed class StaThreadPool : IDisposable, IAsyncDisposable
    {
        private readonly ILogger _logger;
        private readonly Channel<WorkItem> _workChannel;
        private readonly ChannelWriter<WorkItem> _workWriter;
        private readonly ChannelReader<WorkItem> _workReader;
        private readonly CancellationTokenSource _shutdownTokenSource;
        private readonly Thread[] _threads;
        private readonly int _threadCount;
        private readonly string _poolName;
        private volatile bool _disposed;

        /// <summary>
        /// Initializes a new STA thread pool with the specified number of threads.
        /// </summary>
        /// <param name="threadCount">Number of STA threads to create. Defaults to Environment.ProcessorCount.</param>
        /// <param name="poolName">Name for the thread pool (used in logging and thread names).</param>
        public StaThreadPool(int threadCount = 0, string poolName = "StaThreadPool")
        {
            _logger = Log.ForContext<StaThreadPool>();
            _threadCount = threadCount > 0 ? threadCount : Environment.ProcessorCount;
            _poolName = poolName;
            
            _logger.Information("Initializing {PoolName} with {ThreadCount} STA threads", _poolName, _threadCount);

            // Create unbounded channel for work items
            var options = new UnboundedChannelOptions
            {
                SingleReader = false,
                SingleWriter = false,
                AllowSynchronousContinuations = false
            };
            
            _workChannel = Channel.CreateUnbounded<WorkItem>(options);
            _workWriter = _workChannel.Writer;
            _workReader = _workChannel.Reader;
            
            _shutdownTokenSource = new CancellationTokenSource();
            _threads = new Thread[_threadCount];

            // Create and start STA threads with better error handling
            for (int i = 0; i < _threadCount; i++)
            {
                try
                {
                    var thread = new Thread(ThreadWorker)
                    {
                        Name = $"{_poolName}-Thread-{i + 1}",
                        IsBackground = true
                    };
                    
                    // CRITICAL: Set apartment state to STA BEFORE starting the thread
                    thread.SetApartmentState(ApartmentState.STA);
                    _threads[i] = thread;
                    
                    // Start the thread and verify it's running
                    thread.Start(i);
                    
                    // Give the thread a moment to initialize
                    Thread.Sleep(10);
                    
                    _logger.Information("Created STA thread {ThreadName} with apartment state: {ApartmentState}", 
                        thread.Name, thread.GetApartmentState());
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to create STA thread {ThreadIndex}", i);
                    throw new InvalidOperationException($"Failed to create STA thread {i}: {ex.Message}", ex);
                }
            }
            
            _logger.Information("{PoolName} initialized successfully with {ActualThreads} threads", 
                _poolName, _threads.Count(t => t != null));
        }

        /// <summary>
        /// Executes a synchronous function on an STA thread.
        /// </summary>
        /// <typeparam name="T">The return type of the function.</typeparam>
        /// <param name="func">The function to execute.</param>
        /// <param name="timeout">Optional timeout for the operation.</param>
        /// <param name="cancellationToken">Optional cancellation token.</param>
        /// <returns>The result of the function execution.</returns>
        public async Task<T> ExecuteAsync<T>(
            Func<T> func, 
            TimeSpan? timeout = null, 
            CancellationToken cancellationToken = default)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(StaThreadPool));
            
            var tcs = new TaskCompletionSource<T>(TaskCreationOptions.RunContinuationsAsynchronously);
            var workItem = new WorkItem(() =>
            {
                try
                {
                    var result = func();
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            });

            // Set up cancellation
            using var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(
                cancellationToken, _shutdownTokenSource.Token);
            
            combinedCts.Token.Register(() => tcs.TrySetCanceled(combinedCts.Token));

            // Set up timeout if specified
            using var timeoutCts = timeout.HasValue 
                ? new CancellationTokenSource(timeout.Value) 
                : null;
            
            if (timeoutCts != null)
            {
                timeoutCts.Token.Register(() => 
                    tcs.TrySetException(new TimeoutException($"Operation timed out after {timeout}")));
            }

            // Queue the work item
            if (!_workWriter.TryWrite(workItem))
            {
                throw new InvalidOperationException("Failed to queue work item - thread pool may be shutting down");
            }

            return await tcs.Task.ConfigureAwait(false);
        }

        /// <summary>
        /// Executes a synchronous action on an STA thread.
        /// </summary>
        /// <param name="action">The action to execute.</param>
        /// <param name="timeout">Optional timeout for the operation.</param>
        /// <param name="cancellationToken">Optional cancellation token.</param>
        public async Task ExecuteAsync(
            Action action, 
            TimeSpan? timeout = null, 
            CancellationToken cancellationToken = default)
        {
            // Fix: Explicitly call the generic version to avoid infinite recursion
            await ExecuteAsync<bool>(() =>
            {
                action();
                return true; // Return dummy value for consistency
            }, timeout, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Executes an asynchronous function on an STA thread.
        /// Note: The async function will run on the STA thread, but any continuations
        /// may run on different threads unless explicitly marshaled.
        /// </summary>
        /// <typeparam name="T">The return type of the function.</typeparam>
        /// <param name="func">The async function to execute.</param>
        /// <param name="timeout">Optional timeout for the operation.</param>
        /// <param name="cancellationToken">Optional cancellation token.</param>
        /// <returns>The result of the function execution.</returns>
        public async Task<T> ExecuteAsync<T>(
            Func<Task<T>> func, 
            TimeSpan? timeout = null, 
            CancellationToken cancellationToken = default)
        {
            // Fix: Call the synchronous version with a function that waits for the async task
            return await ExecuteAsync(() =>
            {
                // Run the async function synchronously on the STA thread
                return func().GetAwaiter().GetResult();
            }, timeout, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// The main worker loop for each STA thread.
        /// </summary>
        private void ThreadWorker(object? state)
        {
            var threadIndex = (int)state!;
            var threadName = Thread.CurrentThread.Name;
            var currentThread = Thread.CurrentThread;
            
            _logger.Debug("STA thread {ThreadName} starting (apartment state: {ApartmentState})", 
                threadName, currentThread.GetApartmentState());

            try
            {
                // Verify we're running on STA thread - this is CRITICAL for COM operations
                var apartmentState = currentThread.GetApartmentState();
                if (apartmentState != ApartmentState.STA)
                {
                    _logger.Error("CRITICAL: Thread {ThreadName} is {ApartmentState}, not STA! COM operations will fail.", 
                        threadName, apartmentState);
                    return; // Exit the thread worker if not STA
                }

                _logger.Information("STA thread {ThreadName} confirmed STA and ready for COM operations", threadName);

                // THREADING FIX: Use proper async reading pattern to avoid blocking and timeouts
                while (!_shutdownTokenSource.Token.IsCancellationRequested)
                {
                    try
                    {
                        // Use synchronous WaitToRead to avoid complex async context on worker threads
                        var waitTask = _workReader.WaitToReadAsync(_shutdownTokenSource.Token).AsTask();
                        
                        // Wait with shorter timeout to be more responsive to cancellation
                        if (waitTask.Wait(50, _shutdownTokenSource.Token))
                        {
                            if (waitTask.Result && _workReader.TryRead(out var workItem))
                            {
                                // Double-check apartment state before each work item (paranoid but safe)
                                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                                {
                                    _logger.Error("Thread {ThreadName} apartment state changed to {CurrentState}! Skipping work item.", 
                                        threadName, Thread.CurrentThread.GetApartmentState());
                                    continue;
                                }

                                // Execute work item with error handling
                                try
                                {
                                    workItem.Action();
                                }
                                catch (Exception workEx)
                                {
                                    _logger.Error(workEx, "Work item execution failed on thread {ThreadName}", threadName);
                                    // Continue with next work item
                                }
                            }
                        }
                        // If WaitToRead times out, just loop again to check cancellation
                    }
                    catch (OperationCanceledException)
                    {
                        // Expected when shutting down
                        break;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Unhandled exception in STA thread {ThreadName}", threadName);
                        // Sleep briefly to avoid tight error loop
                        Thread.Sleep(100);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.Debug("STA thread {ThreadName} shutting down", threadName);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fatal error in STA thread {ThreadName}", threadName);
            }
            
            _logger.Debug("STA thread {ThreadName} stopped", threadName);
        }

        /// <summary>
        /// Verifies that all threads in the pool are properly configured as STA threads.
        /// This is useful for diagnostics to ensure COM operations will work.
        /// </summary>
        /// <returns>True if all threads are STA, false otherwise.</returns>
        public bool VerifyStaThreads()
        {
            if (_disposed)
                return false;

            var allSta = true;
            for (int i = 0; i < _threadCount; i++)
            {
                var thread = _threads[i];
                if (thread == null || !thread.IsAlive)
                {
                    _logger.Error("Thread {ThreadIndex} is null or not alive", i);
                    allSta = false;
                    continue;
                }

                var apartmentState = thread.GetApartmentState();
                if (apartmentState != ApartmentState.STA)
                {
                    _logger.Error("Thread {ThreadName} has apartment state {ApartmentState}, expected STA", 
                        thread.Name, apartmentState);
                    allSta = false;
                }
                else
                {
                    _logger.Debug("Thread {ThreadName} confirmed STA", thread.Name);
                }
            }

            return allSta;
        }

        /// <summary>
        /// Tests if the thread pool is functional by executing a simple operation.
        /// </summary>
        /// <returns>True if the test operation succeeds, false otherwise.</returns>
        public async Task<bool> TestFunctionality()
        {
            if (_disposed)
                return false;

            try
            {
                var result = await ExecuteAsync(() =>
                {
                    // Simple test - verify we're on STA thread and return success
                    var apartmentState = Thread.CurrentThread.GetApartmentState();
                    return apartmentState == ApartmentState.STA;
                });

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Thread pool functionality test failed");
                return false;
            }
        }

        /// <summary>
        /// Gets the current number of active threads in the pool.
        /// </summary>
        public int ThreadCount => _threadCount;

        /// <summary>
        /// Gets the name of this thread pool.
        /// </summary>
        public string PoolName => _poolName;

        /// <summary>
        /// Gets whether this thread pool has been disposed.
        /// </summary>
        public bool IsDisposed => _disposed;

        /// <summary>
        /// Synchronously disposes the thread pool.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;
            
            _logger.Information("Disposing {PoolName}...", _poolName);
            
            try
            {
                // Signal shutdown and complete the channel
                _shutdownTokenSource.Cancel();
                _workWriter.Complete();

                // Wait for all threads to finish with timeout
                foreach (var thread in _threads)
                {
                    if (!thread.Join(TimeSpan.FromSeconds(10)))
                    {
                        _logger.Warning("Thread {ThreadName} did not shut down gracefully within timeout", thread.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during {PoolName} disposal", _poolName);
            }
            finally
            {
                _shutdownTokenSource.Dispose();
                _disposed = true;
                _logger.Information("{PoolName} disposed", _poolName);
            }
        }

        /// <summary>
        /// Asynchronously disposes the thread pool.
        /// </summary>
        public async ValueTask DisposeAsync()
        {
            if (_disposed) return;
            
            _logger.Information("Disposing {PoolName} asynchronously...", _poolName);
            
            try
            {
                // Signal shutdown and complete the channel
                _shutdownTokenSource.Cancel();
                _workWriter.Complete();

                // Wait for all threads to finish with timeout
                var tasks = new Task[_threadCount];
                for (int i = 0; i < _threadCount; i++)
                {
                    var thread = _threads[i];
                    tasks[i] = Task.Run(() => thread.Join(TimeSpan.FromSeconds(10)));
                }

                await Task.WhenAll(tasks).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during async {PoolName} disposal", _poolName);
            }
            finally
            {
                _shutdownTokenSource.Dispose();
                _disposed = true;
                _logger.Information("{PoolName} disposed asynchronously", _poolName);
            }
        }

        /// <summary>
        /// Represents a work item to be executed on an STA thread.
        /// </summary>
        private readonly struct WorkItem
        {
            public readonly Action Action;

            public WorkItem(Action action)
            {
                Action = action ?? throw new ArgumentNullException(nameof(action));
            }
        }
    }
} 