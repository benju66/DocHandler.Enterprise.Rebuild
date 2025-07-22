using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Base implementation for processing modes
    /// </summary>
    public abstract class ProcessingModeBase : IProcessingMode
    {
        protected readonly ILogger _logger;
        protected IModeContext? _context;
        private bool _disposed = false;

        public abstract string ModeName { get; }
        public abstract string DisplayName { get; }
        public abstract string Description { get; }
        public virtual Version Version => new Version(1, 0, 0);
        public virtual bool IsAvailable => true;

        protected ProcessingModeBase()
        {
            _logger = Log.ForContext(GetType());
        }

        public virtual async Task InitializeAsync(IModeContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger.Information("Initializing mode {ModeName}", ModeName);
            
            try
            {
                await InitializeModeAsync();
                _logger.Information("Mode {ModeName} initialized successfully", ModeName);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to initialize mode {ModeName}", ModeName);
                throw;
            }
        }

        public async Task<ModeProcessingResult> ProcessAsync(ProcessingRequest request, CancellationToken cancellationToken = default)
        {
            if (_context == null)
                throw new InvalidOperationException($"Mode {ModeName} has not been initialized");

            var startTime = DateTime.UtcNow;
            var correlationId = _context.CorrelationId;
            
            _logger.Information("Starting processing for mode {ModeName} with {FileCount} files. CorrelationId: {CorrelationId}", 
                ModeName, request.Files.Count, correlationId);

            try
            {
                // Validate files first
                var validation = ValidateFiles(request.Files);
                if (!validation.IsValid)
                {
                    return new ModeProcessingResult
                    {
                        Success = false,
                        ErrorMessage = validation.ErrorMessage,
                        Duration = DateTime.UtcNow - startTime
                    };
                }

                // Process files
                var result = await ProcessFilesAsync(request, cancellationToken);
                result.Duration = DateTime.UtcNow - startTime;

                _logger.Information("Completed processing for mode {ModeName}. Success: {Success}, Files: {ProcessedCount}/{TotalCount}. CorrelationId: {CorrelationId}",
                    ModeName, result.Success, result.ProcessedFiles.Count(f => f.Success), request.Files.Count, correlationId);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Processing failed for mode {ModeName}. CorrelationId: {CorrelationId}", ModeName, correlationId);
                
                return new ModeProcessingResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    Exception = ex,
                    Duration = DateTime.UtcNow - startTime
                };
            }
        }

        public virtual ValidationResult ValidateFiles(IEnumerable<FileItem> files)
        {
            var fileList = files.ToList();
            var validFiles = new List<FileItem>();
            var invalidFiles = new List<FileItem>();
            var warnings = new List<string>();

            foreach (var file in fileList)
            {
                if (IsFileSupported(file))
                {
                    validFiles.Add(file);
                }
                else
                {
                    invalidFiles.Add(file);
                }
            }

            var isValid = validFiles.Any() && !invalidFiles.Any();
            string? errorMessage = null;

            if (!validFiles.Any())
            {
                errorMessage = $"No files are supported by {DisplayName} mode";
            }
            else if (invalidFiles.Any())
            {
                warnings.Add($"{invalidFiles.Count} files are not supported by {DisplayName} mode");
            }

            return new ValidationResult
            {
                IsValid = isValid,
                ErrorMessage = errorMessage,
                Warnings = warnings,
                ValidFiles = validFiles,
                InvalidFiles = invalidFiles
            };
        }

        public virtual IModeConfiguration GetConfiguration()
        {
            return new ModeConfiguration(ModeName);
        }

        public virtual IModeUIProvider GetUIProvider()
        {
            return new DefaultModeUIProvider(this);
        }

        /// <summary>
        /// Override this method to provide mode-specific initialization logic
        /// </summary>
        protected virtual Task InitializeModeAsync()
        {
            return Task.CompletedTask;
        }

        /// <summary>
        /// Override this method to implement the core processing logic
        /// </summary>
        protected abstract Task<ModeProcessingResult> ProcessFilesAsync(ProcessingRequest request, CancellationToken cancellationToken);

        /// <summary>
        /// Override this method to specify which files are supported by this mode
        /// </summary>
        protected abstract bool IsFileSupported(FileItem file);

        /// <summary>
        /// Get a service from the mode context
        /// </summary>
        protected T GetService<T>() where T : notnull
        {
            if (_context == null)
                throw new InvalidOperationException($"Mode {ModeName} has not been initialized");

            return _context.Services.GetService(typeof(T)) is T service 
                ? service 
                : throw new InvalidOperationException($"Service {typeof(T).Name} is not available");
        }

        /// <summary>
        /// Try to get a service from the mode context
        /// </summary>
        protected T? TryGetService<T>() where T : class
        {
            return _context?.Services.GetService(typeof(T)) as T;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _logger.Information("Disposing mode {ModeName}", ModeName);
                    // Dispose managed resources
                }
                _disposed = true;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }

    /// <summary>
    /// Default implementation of IModeConfiguration
    /// </summary>
    public class ModeConfiguration : IModeConfiguration
    {
        public string ModeName { get; }
        public IDictionary<string, object> Settings { get; } = new Dictionary<string, object>();

        public ModeConfiguration(string modeName)
        {
            ModeName = modeName ?? throw new ArgumentNullException(nameof(modeName));
        }

        public T GetSetting<T>(string key, T defaultValue = default!)
        {
            if (Settings.TryGetValue(key, out var value) && value is T typedValue)
            {
                return typedValue;
            }
            return defaultValue;
        }

        public void SetSetting<T>(string key, T value)
        {
            if (value != null)
            {
                Settings[key] = value;
            }
            else if (Settings.ContainsKey(key))
            {
                Settings.Remove(key);
            }
        }
    }

    /// <summary>
    /// Default implementation of IModeUIProvider
    /// </summary>
    public class DefaultModeUIProvider : IModeUIProvider
    {
        private readonly IProcessingMode _mode;

        public DefaultModeUIProvider(IProcessingMode mode)
        {
            _mode = mode ?? throw new ArgumentNullException(nameof(mode));
        }

        public virtual System.Windows.Controls.UserControl? GetModePanel()
        {
            return null; // Override in specific modes
        }

        public virtual IEnumerable<System.Windows.Controls.MenuItem> GetMenuItems()
        {
            return Enumerable.Empty<System.Windows.Controls.MenuItem>();
        }

        public virtual IEnumerable<object> GetToolBarItems()
        {
            return Enumerable.Empty<object>();
        }

        public virtual void UpdateUIState(object state)
        {
            // Override in specific modes
        }
    }

    /// <summary>
    /// Implementation of IModeContext
    /// </summary>
    public class ModeContext : IModeContext
    {
        public IServiceProvider Services { get; }
        public string CorrelationId { get; }
        public IDictionary<string, object> Properties { get; } = new Dictionary<string, object>();
        public CancellationToken CancellationToken { get; }

        public ModeContext(IServiceProvider services, CancellationToken cancellationToken = default)
        {
            Services = services ?? throw new ArgumentNullException(nameof(services));
            CorrelationId = Guid.NewGuid().ToString();
            CancellationToken = cancellationToken;
        }

        public ModeContext(IServiceProvider services, string correlationId, CancellationToken cancellationToken = default)
        {
            Services = services ?? throw new ArgumentNullException(nameof(services));
            CorrelationId = correlationId ?? throw new ArgumentNullException(nameof(correlationId));
            CancellationToken = cancellationToken;
        }
    }
} 