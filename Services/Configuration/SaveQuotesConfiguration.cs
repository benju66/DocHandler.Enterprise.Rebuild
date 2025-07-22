using System.ComponentModel.DataAnnotations;

namespace DocHandler.Services.Configuration
{
    /// <summary>
    /// Strongly-typed configuration for SaveQuotes mode
    /// </summary>
    public class SaveQuotesConfiguration
    {
        /// <summary>
        /// Whether to automatically scan for company names in files
        /// </summary>
        public bool AutoScanCompanyNames { get; set; } = true;

        /// <summary>
        /// File size limit in MB for .doc files when scanning for company names
        /// </summary>
        [Range(1, 100, ErrorMessage = "File size limit must be between 1 and 100 MB")]
        public int DocFileSizeLimitMB { get; set; } = 10;

        /// <summary>
        /// Whether to scan .doc files for company names (can be slow)
        /// </summary>
        public bool ScanDocFiles { get; set; } = false;

        /// <summary>
        /// Whether to clear the scope selection after processing
        /// </summary>
        public bool ClearScopeAfterProcessing { get; set; } = false;

        /// <summary>
        /// Whether to show recent scopes in the UI
        /// </summary>
        public bool ShowRecentScopes { get; set; } = false;

        /// <summary>
        /// Default scope of work to pre-select
        /// </summary>
        public string DefaultScope { get; set; } = "03-1000";

        /// <summary>
        /// Whether to enable company name detection for this mode
        /// </summary>
        public bool EnableCompanyDetection { get; set; } = true;

        /// <summary>
        /// Maximum number of files to process concurrently in this mode
        /// </summary>
        [Range(1, 10, ErrorMessage = "Max concurrency must be between 1 and 10")]
        public int MaxConcurrency { get; set; } = 3;

        /// <summary>
        /// Timeout for individual file processing in seconds
        /// </summary>
        [Range(30, 600, ErrorMessage = "Timeout must be between 30 and 600 seconds")]
        public int ProcessingTimeoutSeconds { get; set; } = 300;

        /// <summary>
        /// Whether to organize files by company name in subfolders
        /// </summary>
        public bool OrganizeByCompany { get; set; } = true;

        /// <summary>
        /// Whether to organize files by scope of work in subfolders
        /// </summary>
        public bool OrganizeByScope { get; set; } = true;

        /// <summary>
        /// Custom naming pattern for processed files
        /// Available placeholders: {CompanyName}, {Scope}, {OriginalName}, {Date}, {Time}
        /// </summary>
        public string FileNamingPattern { get; set; } = "{CompanyName}_{Scope}_{OriginalName}";

        /// <summary>
        /// Whether to generate processing reports
        /// </summary>
        public bool GenerateReports { get; set; } = true;
    }

    /// <summary>
    /// UI customization settings for SaveQuotes mode
    /// </summary>
    public class SaveQuotesUIConfiguration
    {
        /// <summary>
        /// Whether to show company detection panel
        /// </summary>
        public bool ShowCompanyDetection { get; set; } = true;

        /// <summary>
        /// Whether to show scope selector
        /// </summary>
        public bool ShowScopeSelector { get; set; } = true;

        /// <summary>
        /// Whether to use compact mode (smaller UI elements)
        /// </summary>
        public bool CompactMode { get; set; } = false;

        /// <summary>
        /// Whether to show progress details
        /// </summary>
        public bool ShowProgressDetails { get; set; } = true;

        /// <summary>
        /// Whether to show recent locations panel
        /// </summary>
        public bool ShowRecentLocations { get; set; } = true;

        /// <summary>
        /// Whether to show queue management buttons
        /// </summary>
        public bool ShowQueueManagement { get; set; } = true;
    }
} 