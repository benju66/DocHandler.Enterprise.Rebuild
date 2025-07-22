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
    /// Interface for mode management
    /// </summary>
    public interface IModeManager : IDisposable
    {
        /// <summary>
        /// Current active mode
        /// </summary>
        string? CurrentMode { get; }
        
        /// <summary>
        /// Get all available modes
        /// </summary>
        IEnumerable<ModeDescriptor> GetAvailableModes();
        
        /// <summary>
        /// Switch to a specific mode
        /// </summary>
        Task<bool> SwitchToModeAsync(string modeName);
        
        /// <summary>
        /// Process files using the current mode
        /// </summary>
        Task<ModeProcessingResult> ProcessFilesAsync(IReadOnlyList<FileItem> files, 
            string outputDirectory, 
            IDictionary<string, object> parameters, 
            CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Validate files for the current mode
        /// </summary>
        ValidationResult ValidateFiles(IEnumerable<FileItem> files);
        
        /// <summary>
        /// Check if a mode is available
        /// </summary>
        bool IsModeAvailable(string modeName);
        
        /// <summary>
        /// Event fired when mode changes
        /// </summary>
        event EventHandler<ModeChangedEventArgs>? ModeChanged;
    }

    /// <summary>
    /// Event args for mode change events
    /// </summary>
    public class ModeChangedEventArgs : EventArgs
    {
        public string? PreviousMode { get; set; }
        public string? NewMode { get; set; }
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Mode manager implementation
    /// </summary>
    public class ModeManager : IModeManager
    {
        private readonly IModeRegistry _modeRegistry;
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger _logger;
        private IProcessingMode? _currentModeInstance;
        private bool _disposed = false;

        public string? CurrentMode { get; private set; }

        public event EventHandler<ModeChangedEventArgs>? ModeChanged;

        public ModeManager(IModeRegistry modeRegistry, IServiceProvider serviceProvider)
        {
            _modeRegistry = modeRegistry ?? throw new ArgumentNullException(nameof(modeRegistry));
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _logger = Log.ForContext<ModeManager>();
        }

        public IEnumerable<ModeDescriptor> GetAvailableModes()
        {
            return _modeRegistry.GetAvailableModes();
        }

        public async Task<bool> SwitchToModeAsync(string modeName)
        {
            if (string.IsNullOrWhiteSpace(modeName))
            {
                _logger.Warning("Attempted to switch to null or empty mode name");
                return false;
            }

            var previousMode = CurrentMode;

            try
            {
                _logger.Information("Switching mode from {PreviousMode} to {NewMode}", previousMode, modeName);

                // Check if mode is available
                if (!_modeRegistry.IsRegistered(modeName))
                {
                    _logger.Error("Mode {ModeName} is not registered", modeName);
                    RaiseModeChanged(previousMode, null, false, $"Mode '{modeName}' is not registered");
                    return false;
                }

                // Dispose current mode if any
                if (_currentModeInstance != null)
                {
                    _logger.Debug("Disposing current mode instance: {CurrentMode}", CurrentMode);
                    _currentModeInstance.Dispose();
                    _currentModeInstance = null;
                }

                // Create new mode instance
                _currentModeInstance = await _modeRegistry.CreateModeAsync(modeName, _serviceProvider);
                CurrentMode = modeName;

                _logger.Information("Successfully switched to mode: {ModeName}", modeName);
                RaiseModeChanged(previousMode, modeName, true, null);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to switch to mode: {ModeName}", modeName);
                CurrentMode = null;
                _currentModeInstance = null;
                RaiseModeChanged(previousMode, null, false, ex.Message);
                return false;
            }
        }

        public async Task<ModeProcessingResult> ProcessFilesAsync(IReadOnlyList<FileItem> files, 
            string outputDirectory, 
            IDictionary<string, object> parameters, 
            CancellationToken cancellationToken = default)
        {
            if (_currentModeInstance == null)
            {
                return new ModeProcessingResult
                {
                    Success = false,
                    ErrorMessage = "No mode is currently active"
                };
            }

            var request = new ProcessingRequest
            {
                Files = files,
                OutputDirectory = outputDirectory,
                Parameters = parameters,
                CancellationToken = cancellationToken
            };

            return await _currentModeInstance.ProcessAsync(request, cancellationToken);
        }

        public ValidationResult ValidateFiles(IEnumerable<FileItem> files)
        {
            if (_currentModeInstance == null)
            {
                return new ValidationResult
                {
                    IsValid = false,
                    ErrorMessage = "No mode is currently active"
                };
            }

            return _currentModeInstance.ValidateFiles(files);
        }

        public bool IsModeAvailable(string modeName)
        {
            return !string.IsNullOrWhiteSpace(modeName) && _modeRegistry.IsRegistered(modeName);
        }

        private void RaiseModeChanged(string? previousMode, string? newMode, bool success, string? errorMessage)
        {
            try
            {
                ModeChanged?.Invoke(this, new ModeChangedEventArgs
                {
                    PreviousMode = previousMode,
                    NewMode = newMode,
                    Success = success,
                    ErrorMessage = errorMessage
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error raising ModeChanged event");
            }
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _logger.Information("Disposing ModeManager");
                    
                    if (_currentModeInstance != null)
                    {
                        _currentModeInstance.Dispose();
                        _currentModeInstance = null;
                    }
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
} 