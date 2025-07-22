using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services
{
    // Note: ProcessingRequest and ValidationResult are defined in IProcessingMode.cs



    /// <summary>
    /// Request object for company detection operations
    /// </summary>
    public class CompanyDetectionRequest
    {
        public string FilePath { get; }
        public IProgress<int>? Progress { get; }
        public CancellationToken CancellationToken { get; }

        public CompanyDetectionRequest(string filePath, IProgress<int>? progress = null, 
            CancellationToken cancellationToken = default)
        {
            FilePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
            Progress = progress;
            CancellationToken = cancellationToken;
        }
    }

    /// <summary>
    /// Request object for scope search operations
    /// </summary>
    public class ScopeSearchRequest
    {
        public string SearchTerm { get; }
        public int MaxResults { get; }
        public bool FuzzySearch { get; }

        public ScopeSearchRequest(string searchTerm, int maxResults = 50, bool fuzzySearch = true)
        {
            SearchTerm = searchTerm ?? throw new ArgumentNullException(nameof(searchTerm));
            MaxResults = maxResults;
            FuzzySearch = fuzzySearch;
        }
    }

    /// <summary>
    /// Orchestrates file processing workflows
    /// </summary>
    public interface IFileProcessingOrchestrator
    {
        /// <summary>
        /// Process files using the standard workflow
        /// </summary>
        Task<ModeProcessingResult> ProcessAsync(ProcessingRequest request);

        /// <summary>
        /// Process files in background with progress reporting
        /// </summary>
        Task<ModeProcessingResult> ProcessInBackgroundAsync(ProcessingRequest request, IProgress<string>? progress = null);

        /// <summary>
        /// Process files using Save Quotes workflow
        /// </summary>
        Task<ModeProcessingResult> ProcessSaveQuotesAsync(ProcessingRequest request);

        /// <summary>
        /// Cancel ongoing processing operation
        /// </summary>
        Task CancelProcessingAsync();

        /// <summary>
        /// Check if processing is currently active
        /// </summary>
        bool IsProcessing { get; }
    }

    /// <summary>
    /// Validates files before processing
    /// </summary>
    public interface IFileValidationService
    {
        /// <summary>
        /// Validate a list of files
        /// </summary>
        Task<ValidationResult> ValidateAsync(IEnumerable<FileItem> files, CancellationToken cancellationToken = default);

        /// <summary>
        /// Validate dropped files and convert to FileItems
        /// </summary>
        Task<List<FileItem>> ValidateDroppedFilesAsync(string[] filePaths, CancellationToken cancellationToken = default);

        /// <summary>
        /// Quick validation for UI feedback
        /// </summary>
        ValidationResult ValidateQuick(IEnumerable<FileItem> files);
    }

    /// <summary>
    /// Handles company name detection and scanning
    /// </summary>
    public interface ICompanyDetectionService
    {
        /// <summary>
        /// Scan a file for company names
        /// </summary>
        Task<string?> ScanForCompanyNameAsync(CompanyDetectionRequest request);

        /// <summary>
        /// Scan multiple files for company names
        /// </summary>
        Task<Dictionary<string, string?>> ScanMultipleFilesAsync(IEnumerable<string> filePaths, 
            IProgress<int>? progress = null, CancellationToken cancellationToken = default);

        /// <summary>
        /// Check if auto-scanning is enabled and should be performed
        /// </summary>
        bool ShouldAutoScan(FileItem file);

        /// <summary>
        /// Clear any cached detection results
        /// </summary>
        void ClearDetectionCache();
    }

    /// <summary>
    /// Manages scope of work operations
    /// </summary>
    public interface IScopeManagementService
    {
        /// <summary>
        /// Perform fuzzy search on scopes
        /// </summary>
        Task<List<string>> SearchScopesAsync(ScopeSearchRequest request);

        /// <summary>
        /// Filter scopes based on search term
        /// </summary>
        Task<List<string>> FilterScopesAsync(string searchTerm);

        /// <summary>
        /// Get recent scopes
        /// </summary>
        Task<List<string>> GetRecentScopesAsync();

        /// <summary>
        /// Update recent scope usage
        /// </summary>
        Task UpdateScopeUsageAsync(string scope);

        /// <summary>
        /// Clear recent scopes
        /// </summary>
        Task ClearRecentScopesAsync();
    }




} 