using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services
{
    /// <summary>
    /// Core interface for all processing modes in DocHandler Enterprise
    /// </summary>
    public interface IProcessingMode : IDisposable
    {
        /// <summary>
        /// Unique identifier for this mode
        /// </summary>
        string ModeName { get; }
        
        /// <summary>
        /// Human-readable display name
        /// </summary>
        string DisplayName { get; }
        
        /// <summary>
        /// Description of what this mode does
        /// </summary>
        string Description { get; }
        
        /// <summary>
        /// Version of this mode implementation
        /// </summary>
        Version Version { get; }
        
        /// <summary>
        /// Whether this mode is currently available/enabled
        /// </summary>
        bool IsAvailable { get; }
        
        /// <summary>
        /// Initialize the mode with the given context
        /// </summary>
        Task InitializeAsync(IModeContext context);
        
        /// <summary>
        /// Process files using this mode
        /// </summary>
        Task<ModeProcessingResult> ProcessAsync(ProcessingRequest request, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Validate if the given files can be processed by this mode
        /// </summary>
        ValidationResult ValidateFiles(IEnumerable<FileItem> files);
        
        /// <summary>
        /// Get mode-specific configuration
        /// </summary>
        IModeConfiguration GetConfiguration();
        
        /// <summary>
        /// Get mode-specific UI provider
        /// </summary>
        IModeUIProvider GetUIProvider();
    }

    /// <summary>
    /// Context provided to modes during initialization
    /// </summary>
    public interface IModeContext
    {
        IServiceProvider Services { get; }
        string CorrelationId { get; }
        IDictionary<string, object> Properties { get; }
        CancellationToken CancellationToken { get; }
    }

    /// <summary>
    /// Request for processing files through a mode
    /// </summary>
    public class ProcessingRequest
    {
        public IReadOnlyList<FileItem> Files { get; set; } = new List<FileItem>();
        public string OutputDirectory { get; set; } = string.Empty;
        public IDictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        public CancellationToken CancellationToken { get; set; }
        public IProgress<ProcessingProgress>? Progress { get; set; }
    }

    /// <summary>
    /// Result of processing operation
    /// </summary>
    public class ModeProcessingResult
    {
        public bool Success { get; set; }
        public IReadOnlyList<ProcessedFile> ProcessedFiles { get; set; } = new List<ProcessedFile>();
        public string? ErrorMessage { get; set; }
        public Exception? Exception { get; set; }
        public IDictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
        public TimeSpan Duration { get; set; }
    }

    /// <summary>
    /// Information about a processed file
    /// </summary>
    public class ProcessedFile
    {
        public FileItem OriginalFile { get; set; } = new FileItem();
        public string? OutputPath { get; set; }
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public IDictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// Progress information for processing operations
    /// </summary>
    public class ProcessingProgress
    {
        public int TotalFiles { get; set; }
        public int ProcessedFiles { get; set; }
        public string CurrentFile { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public double PercentComplete => TotalFiles > 0 ? (double)ProcessedFiles / TotalFiles * 100 : 0;
    }

    /// <summary>
    /// Validation result for file compatibility with a mode
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; set; }
        public string? ErrorMessage { get; set; }
        public IList<string> ErrorMessages { get; set; } = new List<string>();
        public IList<string> Messages { get; set; } = new List<string>();
        public IList<string> Warnings { get; set; } = new List<string>();
        public IList<FileItem> ValidFiles { get; set; } = new List<FileItem>();
        public IList<FileItem> InvalidFiles { get; set; } = new List<FileItem>();
    }

    /// <summary>
    /// Mode-specific configuration interface
    /// </summary>
    public interface IModeConfiguration
    {
        string ModeName { get; }
        IDictionary<string, object> Settings { get; }
        T GetSetting<T>(string key, T defaultValue = default!);
        void SetSetting<T>(string key, T value);
    }

    /// <summary>
    /// Mode-specific UI provider interface
    /// </summary>
    public interface IModeUIProvider
    {
        /// <summary>
        /// Get mode-specific UI panel
        /// </summary>
        System.Windows.Controls.UserControl? GetModePanel();
        
        /// <summary>
        /// Get mode-specific menu items
        /// </summary>
        IEnumerable<System.Windows.Controls.MenuItem> GetMenuItems();
        
        /// <summary>
        /// Get mode-specific toolbar items
        /// </summary>
        IEnumerable<object> GetToolBarItems();
        
        /// <summary>
        /// Update UI state based on mode state
        /// </summary>
        void UpdateUIState(object state);
    }
} 