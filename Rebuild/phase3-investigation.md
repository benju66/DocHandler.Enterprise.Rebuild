# Phase 3: Performance & Resilience Investigation Guide
**Duration**: Weeks 11-13
**Objective**: Optimize performance and add enterprise resilience features

## Week 11: Performance Optimization

### Investigation Tasks

#### 1. Performance Baseline Analysis
**Current bottlenecks to investigate**:
- `Services/PerformanceMonitor.cs` - Review current metrics
- Memory allocation patterns
- Thread pool usage
- I/O operations
- COM object lifecycle

**Measurement points**:
```csharp
public class PerformanceBaseline
{
    public double AverageMemoryUsageMB { get; set; }
    public double PeakMemoryUsageMB { get; set; }
    public double AverageProcessingTimeMs { get; set; }
    public int MaxConcurrentOperations { get; set; }
    public double GCPressure { get; set; }
}
```

#### 2. Memory Management Improvements
**Files to optimize**:
- `Services/PdfCacheService.cs` - Implement bounded cache
- `ViewModels/MainViewModel.cs` - Fix event handler leaks
- All services with timers - Ensure proper disposal

**Implement bounded cache**:
```csharp
public class BoundedMemoryCache<TKey, TValue>
{
    private readonly long _maxSizeBytes;
    private readonly IMemoryCache _cache;
    private long _currentSizeBytes;
    
    public async Task<TValue> GetOrAddAsync(
        TKey key, 
        Func<Task<TValue>> factory,
        Func<TValue, long> sizeCalculator)
    {
        // Implement size-aware caching with eviction
    }
}
```

#### 3. Resource Pooling Implementation
**Create pools for**:
```csharp
public interface IResourcePool<T> where T : class
{
    Task<PooledResource<T>> AcquireAsync(CancellationToken cancellationToken = default);
    int Available { get; }
    int InUse { get; }
    PoolStatistics GetStatistics();
}

public class OfficeInstancePool : IResourcePool<IOfficeInstance>
{
    // Pool Word/Excel instances for reuse
    // Implement health checks
    // Auto-recovery from failures
    // Metrics collection
}
```

#### 4. Async I/O Optimization
**Convert synchronous operations**:
```csharp
// ❌ WRONG - Blocking I/O
var config = File.ReadAllText(path);

// ✅ CORRECT - Async I/O
var config = await File.ReadAllTextAsync(path).ConfigureAwait(false);
```

## Week 12: Resilience Patterns

### Investigation Tasks

#### 1. Retry Policy Implementation
**Integrate Polly policies**:
```csharp
public class ModeResiliencePolicies
{
    public IAsyncPolicy<T> GetPolicy<T>(string policyName)
    {
        return policyName switch
        {
            "OfficeConversion" => Policy<T>
                .Handle<COMException>()
                .WaitAndRetryAsync(
                    retryCount: 3,
                    sleepDurationProvider: retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                    onRetry: (outcome, timespan, retryCount, context) =>
                    {
                        var logger = context.Values["Logger"] as ILogger;
                        logger?.Warning("Retry {RetryCount} after {Delay}ms", retryCount, timespan.TotalMilliseconds);
                    }),
                    
            "FileAccess" => Policy<T>
                .Handle<IOException>()
                .WaitAndRetryAsync(5, _ => TimeSpan.FromMilliseconds(100)),
                
            _ => Policy.NoOpAsync<T>()
        };
    }
}
```

#### 2. Circuit Breaker Implementation
**Protect critical resources**:
```csharp
public class OfficeServiceCircuitBreaker
{
    private readonly ICircuitBreaker _circuitBreaker;
    
    public OfficeServiceCircuitBreaker()
    {
        _circuitBreaker = Policy
            .Handle<COMException>()
            .CircuitBreakerAsync(
                handledEventsAllowedBeforeBreaking: 3,
                durationOfBreak: TimeSpan.FromMinutes(1),
                onBreak: (exception, duration) => 
                {
                    // Cleanup Office instances
                    // Log circuit open
                    // Notify monitoring
                },
                onReset: () =>
                {
                    // Reinitialize Office
                    // Log circuit closed
                });
    }
}
```

#### 3. Bulkhead Isolation
**Implement resource isolation**:
```csharp
public class ModeBulkheadPolicy
{
    public static IAsyncPolicy CreateBulkhead(string modeName, ModeResourceLimits limits)
    {
        return Policy.BulkheadAsync(
            maxParallelization: limits.MaxConcurrency,
            maxQueuingActions: limits.MaxQueueSize,
            onBulkheadRejectedAsync: async context =>
            {
                var logger = context.Values["Logger"] as ILogger;
                logger?.Warning("Bulkhead rejected operation for mode {Mode}", modeName);
            });
    }
}
```

#### 4. Health Monitoring System
**Implement comprehensive health checks**:
```csharp
public interface IModeHealthCheck
{
    string Name { get; }
    Task<HealthCheckResult> CheckHealthAsync(CancellationToken cancellationToken = default);
}

public class OfficeHealthCheck : IModeHealthCheck
{
    public async Task<HealthCheckResult> CheckHealthAsync(CancellationToken cancellationToken)
    {
        try
        {
            // Test Office connectivity
            // Check COM object count
            // Verify memory usage
            // Test file operations
            
            return HealthCheckResult.Healthy("Office services operational");
        }
        catch (Exception ex)
        {
            return HealthCheckResult.Unhealthy("Office services failing", ex);
        }
    }
}
```

## Week 13: Scalability Features

### Investigation Tasks

#### 1. Queue Management Redesign
**Current issues in `SaveQuotesQueueService.cs`**:
- No priority handling
- Limited concurrency control
- No load distribution

**Implement enhanced queue**:
```csharp
public interface IScalableQueueService<T>
{
    Task EnqueueAsync(T item, QueuePriority priority = QueuePriority.Normal);
    Task<IEnumerable<T>> DequeueBatchAsync(int maxItems, CancellationToken cancellationToken);
    QueueStatistics GetStatistics();
    void SetConcurrencyLimit(int limit);
}

public class ModeAwareQueueService : IScalableQueueService<ProcessingRequest>
{
    private readonly Dictionary<string, PriorityQueue<ProcessingRequest>> _modeQueues;
    private readonly ILoadBalancer _loadBalancer;
    
    public async Task ProcessAsync()
    {
        // Distribute work based on:
        // - Mode priority
        // - Resource availability
        // - Current load
        // - SLA requirements
    }
}
```

#### 2. State Management
**Move from in-memory to persistent**:
```csharp
public interface IStateManager
{
    Task SaveStateAsync<T>(string key, T state) where T : class;
    Task<T> LoadStateAsync<T>(string key) where T : class;
    Task<bool> ExistsAsync(string key);
    Task DeleteStateAsync(string key);
}

public class PersistentStateManager : IStateManager
{
    // Options: SQLite, Redis, or file-based
    // Include versioning
    // Support migrations
    // Enable querying
}
```

#### 3. Message Bus Integration
**For distributed scenarios**:
```csharp
public interface IMessageBus
{
    Task PublishAsync<T>(T message) where T : class;
    Task<IDisposable> SubscribeAsync<T>(Func<T, Task> handler) where T : class;
}

public class ModeEventBus : IMessageBus
{
    // Events to support:
    // - ModeStarted
    // - ModeCompleted
    // - ModeFailed
    // - ResourceExhausted
    // - HealthCheckFailed
}
```

## Performance Testing Suite

### Load Tests
```csharp
[Test]
public async Task LoadTest_ConcurrentOperations()
{
    // Arrange
    var tasks = new List<Task<ProcessingResult>>();
    var fileCount = 100;
    var concurrency = 10;
    
    // Act
    using (var semaphore = new SemaphoreSlim(concurrency))
    {
        for (int i = 0; i < fileCount; i++)
        {
            await semaphore.WaitAsync();
            tasks.Add(ProcessFileAsync().ContinueWith(t => 
            {
                semaphore.Release();
                return t.Result;
            }));
        }
    }
    
    var results = await Task.WhenAll(tasks);
    
    // Assert
    Assert.That(results.All(r => r.Success));
    Assert.That(MemoryUsage, Is.LessThan(500_000_000)); // 500MB
}
```

### Stress Tests
```csharp
[Test]
public async Task StressTest_MemoryPressure()
{
    // Continuously process files while monitoring memory
    // Verify garbage collection works
    // Ensure no memory leaks
    // Check performance degradation
}
```

## Monitoring Implementation

### Metrics to Collect
```csharp
public class ModeMetrics
{
    // Performance
    public double ProcessingTimeP50 { get; set; }
    public double ProcessingTimeP95 { get; set; }
    public double ProcessingTimeP99 { get; set; }
    
    // Resource Usage
    public double MemoryUsageMB { get; set; }
    public int ThreadCount { get; set; }
    public int ComObjectCount { get; set; }
    
    // Business Metrics
    public int FilesProcessed { get; set; }
    public int FailureCount { get; set; }
    public double SuccessRate { get; set; }
    
    // Health
    public int CircuitBreakerOpenCount { get; set; }
    public int BulkheadRejectionCount { get; set; }
}
```

## Deliverables

1. **Performance Improvements**
   - Memory usage reduced by 50%
   - Processing time improved by 30%
   - Resource pooling implemented
   - Async I/O throughout

2. **Resilience Features**
   - Retry policies for all operations
   - Circuit breakers protecting resources
   - Bulkhead isolation per mode
   - Health monitoring system

3. **Scalability Enhancements**
   - Priority queue system
   - Persistent state management
   - Message bus ready
   - Load distribution capable

4. **Testing & Monitoring**
   - Load test suite
   - Stress test suite
   - Performance benchmarks
   - Monitoring dashboard

## Success Criteria
- [ ] Memory usage < 300MB under normal load
- [ ] Can process 100 files concurrently
- [ ] 99% success rate with retries
- [ ] Circuit breakers prevent cascading failures
- [ ] Health checks detect issues < 30 seconds
- [ ] Queue handles 1000 items without issues
- [ ] Performance metrics collected continuously
- [ ] Zero memory leaks in 24-hour test