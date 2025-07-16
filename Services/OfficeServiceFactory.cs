using System;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Serilog;

namespace DocHandler.Services
{
    public class OfficeServiceFactory : IOfficeServiceFactory, IDisposable
    {
        private readonly ILogger _logger;
        private readonly IConfigurationService? _configService;
        private readonly IProcessManager? _processManager;
        private readonly SemaphoreSlim _initializationSemaphore;
        private readonly object _lockObject = new object();
        private bool _disposed;
        private bool? _officeAvailable;
        private DateTime _lastOfficeCheck = DateTime.MinValue;
        private readonly TimeSpan _officeCheckCacheTimeout = TimeSpan.FromMinutes(5);

        public OfficeServiceFactory(IConfigurationService? configService = null, IProcessManager? processManager = null)
        {
            _logger = Log.ForContext<OfficeServiceFactory>();
            _configService = configService;
            _processManager = processManager;
            _initializationSemaphore = new SemaphoreSlim(1, 1);
            
            _logger.Information("Office Service Factory initialized");
        }

        public async Task<bool> IsOfficeAvailableAsync()
        {
            // Use cached result if recent
            if (_officeAvailable.HasValue && 
                DateTime.Now - _lastOfficeCheck < _officeCheckCacheTimeout)
            {
                return _officeAvailable.Value;
            }

            return await Task.Run(() => CheckOfficeAvailability());
        }

        private bool CheckOfficeAvailability()
        {
            try
            {
                _logger.Debug("Checking Office availability...");
                
                // Check for Word
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    _logger.Warning("Word.Application ProgID not found");
                    _officeAvailable = false;
                    _lastOfficeCheck = DateTime.Now;
                    return false;
                }

                // Check for Excel
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    _logger.Warning("Excel.Application ProgID not found");
                    _officeAvailable = false;
                    _lastOfficeCheck = DateTime.Now;
                    return false;
                }

                _logger.Information("Microsoft Office is available");
                _officeAvailable = true;
                _lastOfficeCheck = DateTime.Now;
                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error checking Office availability");
                _officeAvailable = false;
                _lastOfficeCheck = DateTime.Now;
                return false;
            }
        }

        public async Task<IOfficeConversionService> CreateOfficeServiceAsync()
        {
            if (!await IsOfficeAvailableAsync())
            {
                throw new InvalidOperationException("Microsoft Office is not available on this system");
            }

            await _initializationSemaphore.WaitAsync();
            try
            {
                _logger.Information("Creating Office conversion service...");
                
                // Use the robust service that doesn't have recursion issues
                var service = new RobustOfficeConversionService(maxConcurrency: 2);
                return new RobustOfficeConversionServiceWrapper(service);
            }
            finally
            {
                _initializationSemaphore.Release();
            }
        }

        public async Task<ISessionAwareOfficeService> CreateSessionOfficeServiceAsync()
        {
            if (!await IsOfficeAvailableAsync())
            {
                throw new InvalidOperationException("Microsoft Office is not available on this system");
            }

            await _initializationSemaphore.WaitAsync();
            try
            {
                _logger.Information("Creating session-aware Office service...");
                
                // Create service directly (WPF main thread is already STA)
                var service = new SessionAwareOfficeService();
                
                return new SessionAwareOfficeServiceWrapper(service);
            }
            finally
            {
                _initializationSemaphore.Release();
            }
        }

        public async Task<ISessionAwareExcelService> CreateSessionExcelServiceAsync()
        {
            if (!await IsOfficeAvailableAsync())
            {
                throw new InvalidOperationException("Microsoft Office is not available on this system");
            }

            await _initializationSemaphore.WaitAsync();
            try
            {
                _logger.Information("Creating session-aware Excel service...");
                
                // Create service directly (WPF main thread is already STA)
                var service = new SessionAwareExcelService();
                
                return new SessionAwareExcelServiceWrapper(service);
            }
            finally
            {
                _initializationSemaphore.Release();
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _initializationSemaphore?.Dispose();
                _disposed = true;
                _logger.Information("Office Service Factory disposed");
            }
        }
    }

    // Wrapper implementations to make existing services implement interfaces
    public class RobustOfficeConversionServiceWrapper : IOfficeConversionService
    {
        private readonly RobustOfficeConversionService _inner;

        public RobustOfficeConversionServiceWrapper(RobustOfficeConversionService inner)
        {
            _inner = inner;
        }

        public async Task<string?> ConvertToPdfAsync(string inputPath, string outputPath)
        {
            var result = await _inner.ConvertWordToPdf(inputPath, outputPath);
            return result.Success ? result.OutputPath : null;
        }

        public async Task<bool> IsOfficeAvailableAsync()
        {
            return await Task.FromResult(_inner.IsOfficeInstalled());
        }

        public void Cleanup()
        {
            // Cleanup is handled internally by the robust service
        }

        public void Dispose()
        {
            _inner?.Dispose();
        }
    }

    public class SessionAwareOfficeServiceWrapper : ISessionAwareOfficeService
    {
        internal readonly SessionAwareOfficeService _inner;

        public SessionAwareOfficeServiceWrapper(SessionAwareOfficeService inner)
        {
            _inner = inner;
        }

        public async Task<string?> ConvertToPdfAsync(string inputPath, string outputPath)
        {
            var result = await _inner.ConvertWordToPdf(inputPath, outputPath);
            return result.Success ? result.OutputPath : null;
        }

        public async Task<string?> ExtractTextAsync(string filePath)
        {
            // SessionAwareOfficeService doesn't have text extraction - return null for now
            return await Task.FromResult<string?>(null);
        }

        public async Task<bool> IsAvailableAsync()
        {
            return await Task.FromResult(_inner.IsOfficeInstalled());
        }

        public void Dispose()
        {
            _inner?.Dispose();
        }
    }

    public class SessionAwareExcelServiceWrapper : ISessionAwareExcelService
    {
        internal readonly SessionAwareExcelService _inner;

        public SessionAwareExcelServiceWrapper(SessionAwareExcelService inner)
        {
            _inner = inner;
        }

        public async Task<string?> ConvertToPdfAsync(string inputPath, string outputPath)
        {
            var result = await _inner.ConvertSpreadsheetToPdf(inputPath, outputPath);
            return result.Success ? result.OutputPath : null;
        }

        public async Task<string?> ExtractTextAsync(string filePath)
        {
            // SessionAwareExcelService doesn't have text extraction - return null for now
            return await Task.FromResult<string?>(null);
        }

        public async Task<bool> IsAvailableAsync()
        {
            return await Task.FromResult(true); // Excel service doesn't have health check method
        }

        public void Dispose()
        {
            _inner?.Dispose();
        }
    }
} 