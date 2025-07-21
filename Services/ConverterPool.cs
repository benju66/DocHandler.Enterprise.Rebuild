using System;
using System.Collections.Concurrent;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Pool of Office converters for parallel document processing.
    /// Automatically sizes based on CPU cores and adjusts based on memory pressure.
    /// </summary>
    public class ConverterPool : IDisposable
    {
        private readonly ILogger _logger = Log.ForContext<ConverterPool>();
        private readonly ConcurrentBag<PooledConverter> _availableConverters = new();
        private readonly ConcurrentDictionary<int, PooledConverter> _allConverters = new();
        private readonly SemaphoreSlim _poolSemaphore;
        private readonly StaThreadPool _staThreadPool;
        private readonly object _sizeLock = new object();
        private int _currentPoolSize;
        private int _maxPoolSize;
        private bool _disposed;
        private int _nextConverterId = 0;

        public int CurrentPoolSize => _currentPoolSize;
        public int AvailableConverters => _availableConverters.Count;

        public ConverterPool(int maxPoolSize = 0)
        {
            // Automatic sizing based on CPU cores
            _maxPoolSize = maxPoolSize > 0 ? maxPoolSize : CalculateOptimalPoolSize();
            _currentPoolSize = _maxPoolSize;
            _poolSemaphore = new SemaphoreSlim(_maxPoolSize, _maxPoolSize);
            _staThreadPool = new StaThreadPool(_maxPoolSize, "ConverterPool");
            
            _logger.Information("Converter pool initialized with size: {Size} (CPU cores: {Cores})", 
                _maxPoolSize, Environment.ProcessorCount);
        }

        private int CalculateOptimalPoolSize()
        {
            return Environment.ProcessorCount switch
            {
                <= 2 => 1,  // Dual-core: Single converter
                <= 4 => 2,  // Quad-core: 2 converters
                <= 8 => 3,  // 6-8 cores: 3 converters  
                _ => 4      // 8+ cores: Max 4 converters
            };
        }

        public async Task<PooledConverter> RentConverterAsync(CancellationToken cancellationToken = default)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(ConverterPool));

            await _poolSemaphore.WaitAsync(cancellationToken);
            
            try
            {
                // Check current pool size
                bool shouldCreateNew = false;
                lock (_sizeLock)
                {
                    if (_allConverters.Count > _currentPoolSize)
                    {
                        // Pool was reduced, don't create new converters
                        shouldCreateNew = false;
                    }
                    else
                    {
                        shouldCreateNew = true;
                    }
                }

                if (!shouldCreateNew)
                {
                    _poolSemaphore.Release();
                    return await RentConverterAsync(cancellationToken);
                }

                // Try to get existing converter
                if (_availableConverters.TryTake(out var converter))
                {
                    if (!converter.IsExpired && !converter.IsDisposed)
                    {
                        _logger.Debug("Renting existing converter {Id}", converter.Id);
                        return converter;
                    }
                    
                    // Converter expired or disposed, remove from tracking
                    _allConverters.TryRemove(converter.Id, out _);
                    converter.Dispose();
                }
                
                // Create new converter on STA thread
                var id = Interlocked.Increment(ref _nextConverterId);
                var newConverter = await _staThreadPool.ExecuteAsync(() => 
                {
                    _logger.Debug("Creating new converter {Id} on STA thread", id);
                    return new PooledConverter(this, id);
                });
                
                _allConverters.TryAdd(id, newConverter);
                return newConverter;
            }
            catch
            {
                _poolSemaphore.Release();
                throw;
            }
        }
        
        internal void ReturnConverter(PooledConverter converter)
        {
            if (_disposed || converter.IsDisposed)
            {
                _allConverters.TryRemove(converter.Id, out _);
                converter.Dispose();
                _poolSemaphore.Release();
                return;
            }

            lock (_sizeLock)
            {
                // Check if pool size was reduced
                if (_allConverters.Count > _currentPoolSize || converter.IsExpired)
                {
                    _logger.Debug("Disposing converter {Id} (pool size reduced or expired)", converter.Id);
                    _allConverters.TryRemove(converter.Id, out _);
                    converter.Dispose();
                    _poolSemaphore.Release();
                    return;
                }
            }

            _logger.Debug("Returning converter {Id} to pool", converter.Id);
            _availableConverters.Add(converter);
            _poolSemaphore.Release();
        }
        
        public void AdjustPoolSize(int newSize)
        {
            lock (_sizeLock)
            {
                var oldSize = _currentPoolSize;
                _currentPoolSize = Math.Max(1, Math.Min(newSize, _maxPoolSize));
                
                if (_currentPoolSize < oldSize)
                {
                    _logger.Information("Reducing pool size from {OldSize} to {NewSize}", 
                        oldSize, _currentPoolSize);
                    
                    // Remove excess converters
                    var toRemove = oldSize - _currentPoolSize;
                    for (int i = 0; i < toRemove && _availableConverters.TryTake(out var converter); i++)
                    {
                        _allConverters.TryRemove(converter.Id, out _);
                        converter.Dispose();
                    }
                }
                else if (_currentPoolSize > oldSize)
                {
                    _logger.Information("Increasing pool size from {OldSize} to {NewSize}", 
                        oldSize, _currentPoolSize);
                }
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            _logger.Information("Disposing converter pool");
            
            // Dispose all converters
            foreach (var converter in _allConverters.Values)
            {
                try
                {
                    converter.Dispose();
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing converter {Id}", converter.Id);
                }
            }
            
            _allConverters.Clear();
            _availableConverters.Clear();
            
            _poolSemaphore?.Dispose();
            _staThreadPool?.Dispose();
        }
    }
    
    public class PooledConverter : IDisposable
    {
        private readonly ConverterPool _pool;
        private readonly ReliableOfficeConverter _converter;
        private readonly ILogger _logger;
        private int _useCount;
        private readonly int _maxUses = 20;
        private bool _disposed;
        private bool _rented;
        
        public int Id { get; }
        public bool IsExpired => _useCount >= _maxUses;
        public bool IsDisposed => _disposed;
        
        internal PooledConverter(ConverterPool pool, int id)
        {
            _pool = pool;
            Id = id;
            _logger = Log.ForContext<PooledConverter>().ForContext("ConverterId", id);
            _converter = new ReliableOfficeConverter();
            _rented = true;
        }
        
        public ConversionResult ConvertWordToPdf(string input, string output, Action<string> progressCallback = null)
        {
            if (_disposed) throw new ObjectDisposedException($"PooledConverter[{Id}]");
            if (!_rented) throw new InvalidOperationException($"Converter {Id} is not rented");
            
            _useCount++;
            progressCallback?.Invoke($"Converting {System.IO.Path.GetFileName(input)}...");
            
            try
            {
                return _converter.ConvertWordToPdf(input, output);
            }
            finally
            {
                progressCallback?.Invoke($"Completed {System.IO.Path.GetFileName(input)}");
            }
        }
        
        public ConversionResult ConvertExcelToPdf(string input, string output, Action<string> progressCallback = null)
        {
            if (_disposed) throw new ObjectDisposedException($"PooledConverter[{Id}]");
            if (!_rented) throw new InvalidOperationException($"Converter {Id} is not rented");
            
            _useCount++;
            progressCallback?.Invoke($"Converting {System.IO.Path.GetFileName(input)}...");
            
            try
            {
                return _converter.ConvertExcelToPdf(input, output);
            }
            finally
            {
                progressCallback?.Invoke($"Completed {System.IO.Path.GetFileName(input)}");
            }
        }
        
        public void Return()
        {
            if (!_rented) return;
            _rented = false;
            _pool.ReturnConverter(this);
        }
        
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _rented = false;
            
            try
            {
                _converter?.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error disposing underlying converter");
            }
        }
    }
} 