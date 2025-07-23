using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services
{
    // Backward compatibility aliases for existing implementations
    public interface ICompanyDetectionService : IEnhancedCompanyDetectionService
    {
        // Legacy method signatures for compatibility
        Task<string?> DetectCompanyAsync(CompanyDetectionRequest request);
        Task<string?> ScanForCompanyNameAsync(CompanyDetectionRequest request);
    }

    public interface IFileValidationService : IEnhancedFileValidationService  
    {
        // Legacy method signatures for compatibility
        Task<LegacyValidationResult> ValidateAsync(IEnumerable<FileItem> files, CancellationToken cancellationToken = default);
        Task<List<FileItem>> ValidateDroppedFilesAsync(string[] filePaths, CancellationToken cancellationToken = default);
        LegacyValidationResult ValidateQuick(IEnumerable<FileItem> files);
    }

    // Legacy data model for compatibility
    public class CompanyDetectionRequest
    {
        public string FilePath { get; set; } = "";
        public string? CompanyHint { get; set; }
        public bool UseCache { get; set; } = true;
        public int TimeoutSeconds { get; set; } = 30;
        public Dictionary<string, object> Options { get; set; } = new();
    }

    public class LegacyValidationResult
    {
        public bool IsValid { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> Warnings { get; set; } = new();
        public List<FileItem> ValidFiles { get; set; } = new();
        public List<FileItem> InvalidFiles { get; set; } = new();
    }

    /// <summary>
    /// Enhanced file validation service with security and performance checks
    /// </summary>
    public interface IEnhancedFileValidationService
    {
        /// <summary>
        /// Validates a single file with comprehensive security and format checks
        /// </summary>
        Task<EnhancedFileValidationResult> ValidateFileAsync(string filePath, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Validates multiple files in parallel
        /// </summary>
        Task<List<EnhancedFileValidationResult>> ValidateFilesAsync(IEnumerable<string> filePaths, IProgress<ValidationProgress>? progress = null, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Checks if file type is supported
        /// </summary>
        bool IsFileTypeSupported(string filePath);
        
        /// <summary>
        /// Gets security risk assessment for a file
        /// </summary>
        Task<SecurityRiskAssessment> AssessSecurityRiskAsync(string filePath);
        
        /// <summary>
        /// Sanitizes file paths and names
        /// </summary>
        string SanitizeFilePath(string filePath);
    }

    /// <summary>
    /// Advanced company detection service with caching and AI-enhanced detection
    /// </summary>
    public interface IEnhancedCompanyDetectionService
    {
        /// <summary>
        /// Detects company name from document with confidence scoring
        /// </summary>
        Task<EnhancedCompanyDetectionResult> DetectCompanyAsync(string filePath, IProgress<int>? progress = null, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Batch company detection for multiple files
        /// </summary>
        Task<List<EnhancedCompanyDetectionResult>> DetectCompaniesAsync(IEnumerable<string> filePaths, IProgress<BatchProgress>? progress = null, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Validates detected company name against known companies
        /// </summary>
        Task<CompanyValidationResult> ValidateCompanyAsync(string companyName);
        
        /// <summary>
        /// Gets suggestions for similar company names
        /// </summary>
        Task<List<string>> GetCompanySuggestionsAsync(string partialName, int maxSuggestions = 10);
        
        /// <summary>
        /// Adds a new company to the knowledge base
        /// </summary>
        Task<bool> AddCompanyAsync(string companyName, List<string>? aliases = null);
        
        /// <summary>
        /// Updates company usage statistics
        /// </summary>
        Task IncrementCompanyUsageAsync(string companyName);
    }

    /// <summary>
    /// Scope of work management service with fuzzy matching and learning
    /// </summary>
    public interface IScopeManagementService
    {
        /// <summary>
        /// Gets all available scopes
        /// </summary>
        Task<List<ScopeInfo>> GetAllScopesAsync();
        
        /// <summary>
        /// Searches scopes with fuzzy matching
        /// </summary>
        Task<List<ScopeInfo>> SearchScopesAsync(string searchTerm, double minConfidence = 0.7);
        
        /// <summary>
        /// Detects scope from document content
        /// </summary>
        Task<ScopeDetectionResult> DetectScopeAsync(string filePath, string? companyName = null);
        
        /// <summary>
        /// Adds a new scope to the system
        /// </summary>
        Task<bool> AddScopeAsync(string scopeName, string? description = null, List<string>? keywords = null);
        
        /// <summary>
        /// Updates scope usage statistics
        /// </summary>
        Task IncrementScopeUsageAsync(string scopeName);
        
        /// <summary>
        /// Gets most frequently used scopes
        /// </summary>
        Task<List<ScopeInfo>> GetMostUsedScopesAsync(int count = 10);
        
        /// <summary>
        /// Learns scope patterns from successful processing
        /// </summary>
        Task LearnScopePatternAsync(string filePath, string scopeName, string? companyName = null);

        // Legacy methods for compatibility
        Task<List<ScopeInfo>> GetRecentScopesAsync(int count = 10);
        Task<List<ScopeInfo>> FilterScopesAsync(string searchTerm);
        Task<bool> ValidateScopeAsync(string scopeName);
    }

    /// <summary>
    /// Document processing workflow orchestration service
    /// </summary>
    public interface IDocumentWorkflowService
    {
        /// <summary>
        /// Processes a single document through the complete workflow
        /// </summary>
        Task<DocumentProcessingResult> ProcessDocumentAsync(DocumentProcessingRequest request, IProgress<WorkflowProgress>? progress = null, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Processes multiple documents in parallel
        /// </summary>
        Task<BatchProcessingResult> ProcessDocumentsAsync(List<DocumentProcessingRequest> requests, BatchProcessingOptions? options = null, IProgress<BatchProgress>? progress = null, CancellationToken cancellationToken = default);
        
        /// <summary>
        /// Gets processing status for a document
        /// </summary>
        Task<WorkflowStatus> GetProcessingStatusAsync(string correlationId);
        
        /// <summary>
        /// Cancels processing for a specific document
        /// </summary>
        Task<bool> CancelProcessingAsync(string correlationId);
    }

    // Enhanced Data Models for Business Services (avoiding conflicts with existing classes)

    public class EnhancedFileValidationResult
    {
        public string FilePath { get; set; } = "";
        public bool IsValid { get; set; }
        public bool IsSecure { get; set; }
        public List<string> Errors { get; set; } = new();
        public List<string> Warnings { get; set; } = new();
        public SecurityRiskLevel RiskLevel { get; set; }
        public long FileSizeBytes { get; set; }
        public string FileType { get; set; } = "";
        public DateTime ValidationTime { get; set; }
        public TimeSpan ValidationDuration { get; set; }
    }

    public class EnhancedCompanyDetectionResult
    {
        public string FilePath { get; set; } = "";
        public string? DetectedCompany { get; set; }
        public double Confidence { get; set; }
        public List<CompanyMatch> AlternativeMatches { get; set; } = new();
        public TimeSpan ProcessingTime { get; set; }
        public string? ErrorMessage { get; set; }
        public Dictionary<string, object> Metadata { get; set; } = new();
    }

    public class CompanyMatch
    {
        public string CompanyName { get; set; } = "";
        public double Confidence { get; set; }
        public string Source { get; set; } = "";
        public List<string> MatchedKeywords { get; set; } = new();
    }

    public class CompanyValidationResult
    {
        public string CompanyName { get; set; } = "";
        public bool IsKnownCompany { get; set; }
        public bool IsValid { get; set; }
        public List<string> Suggestions { get; set; } = new();
        public int UsageCount { get; set; }
        public DateTime LastUsed { get; set; }
    }

    public class ScopeDetectionResult
    {
        public string FilePath { get; set; } = "";
        public string? DetectedScope { get; set; }
        public double Confidence { get; set; }
        public List<ScopeMatch> AlternativeMatches { get; set; } = new();
        public TimeSpan ProcessingTime { get; set; }
        public string? ErrorMessage { get; set; }
    }

    public class ScopeMatch
    {
        public string ScopeName { get; set; } = "";
        public double Confidence { get; set; }
        public List<string> MatchedKeywords { get; set; } = new();
        public string? Description { get; set; }
    }

    public class ScopeInfo
    {
        public string Name { get; set; } = "";
        public string? Description { get; set; }
        public List<string> Keywords { get; set; } = new();
        public int UsageCount { get; set; }
        public DateTime LastUsed { get; set; }
        public DateTime CreatedAt { get; set; }
        public bool IsActive { get; set; } = true;
    }

    public class DocumentProcessingRequest
    {
        public string CorrelationId { get; set; } = Guid.NewGuid().ToString();
        public string FilePath { get; set; } = "";
        public string OutputPath { get; set; } = "";
        public string? CompanyName { get; set; }
        public string? ScopeName { get; set; }
        public Dictionary<string, object> Options { get; set; } = new();
        public DateTime RequestTime { get; set; } = DateTime.UtcNow;
    }

    public class DocumentProcessingResult
    {
        public string CorrelationId { get; set; } = "";
        public string FilePath { get; set; } = "";
        public string? OutputPath { get; set; }
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> Warnings { get; set; } = new();
        public TimeSpan ProcessingTime { get; set; }
        public ProcessingMetrics Metrics { get; set; } = new();
        public DateTime CompletedAt { get; set; }
    }

    public class BatchProcessingResult
    {
        public int TotalFiles { get; set; }
        public int SuccessfulFiles { get; set; }
        public int FailedFiles { get; set; }
        public List<DocumentProcessingResult> Results { get; set; } = new();
        public TimeSpan TotalProcessingTime { get; set; }
        public ProcessingMetrics AggregateMetrics { get; set; } = new();
    }

    public class BatchProcessingOptions
    {
        public int MaxConcurrency { get; set; } = Environment.ProcessorCount;
        public TimeSpan Timeout { get; set; } = TimeSpan.FromMinutes(30);
        public bool StopOnError { get; set; } = false;
        public bool EnableProgressReporting { get; set; } = true;
        public string? OutputDirectory { get; set; }
    }

    public class ProcessingMetrics
    {
        public long MemoryUsedBytes { get; set; }
        public TimeSpan CpuTime { get; set; }
        public int FilesProcessed { get; set; }
        public long BytesProcessed { get; set; }
        public int ErrorCount { get; set; }
        public int WarningCount { get; set; }
    }

    public class UIState
    {
        public bool IsProcessing { get; set; }
        public double ProgressValue { get; set; }
        public string StatusMessage { get; set; } = "";
        public bool CanProcess { get; set; }
        public string ProcessButtonText { get; set; } = "Process Files";
        public bool SaveQuotesMode { get; set; }
        public string? SelectedCompany { get; set; }
        public string? SelectedScope { get; set; }
        public Dictionary<string, object> CustomProperties { get; set; } = new();
    }

    public class UIStateUpdate
    {
        public bool? IsProcessing { get; set; }
        public double? ProgressValue { get; set; }
        public string? StatusMessage { get; set; }
        public bool? CanProcess { get; set; }
        public string? ProcessButtonText { get; set; }
        public Dictionary<string, object>? CustomProperties { get; set; }
    }

    public class ValidationProgress
    {
        public int TotalFiles { get; set; }
        public int CompletedFiles { get; set; }
        public string? CurrentFile { get; set; }
        public double PercentComplete => TotalFiles > 0 ? (double)CompletedFiles / TotalFiles * 100 : 0;
    }

    public class BatchProgress
    {
        public int TotalItems { get; set; }
        public int CompletedItems { get; set; }
        public int FailedItems { get; set; }
        public string? CurrentItem { get; set; }
        public double PercentComplete => TotalItems > 0 ? (double)CompletedItems / TotalItems * 100 : 0;
    }

    public class WorkflowProgress
    {
        public string? CurrentOperation { get; set; }
        public double PercentComplete { get; set; }
        public string? Status { get; set; }
        public TimeSpan Elapsed { get; set; }
        public TimeSpan? EstimatedRemaining { get; set; }
    }

    public enum SecurityRiskLevel
    {
        None = 0,
        Low = 1,
        Medium = 2,
        High = 3,
        Critical = 4
    }

    public class SecurityRiskAssessment
    {
        public SecurityRiskLevel RiskLevel { get; set; }
        public List<string> RiskFactors { get; set; } = new();
        public string? Recommendation { get; set; }
        public bool IsBlocked { get; set; }
    }

    public enum WorkflowStatus
    {
        Pending,
        Validating,
        Processing,
        Completed,
        Failed,
        Cancelled
    }

    // UI State Management Service
    public interface IUIStateService
    {
        // Progress Management
        Task UpdateProgressAsync(double progressValue, string? statusMessage = null);
        Task ResetProgressAsync();
        Task<double> GetCurrentProgressAsync();

        // Status Management  
        Task UpdateStatusAsync(string message);
        Task UpdateQueueStatusAsync(string message);
        Task<string> GetCurrentStatusAsync();

        // Processing State Management
        Task SetProcessingAsync(bool isProcessing);
        Task<bool> IsProcessingAsync();

        // UI Synchronization
        Task InvokeOnUIThreadAsync(Action action);
        Task<T> InvokeOnUIThreadAsync<T>(Func<T> function);

        // Complex UI State Updates
        Task RefreshUIStateAsync(UIStateContext context);
        Task UpdateCanProcessStateAsync(bool canProcess, string? buttonText = null);

        // Error Display
        Task ShowErrorAsync(string title, string message);
        Task ShowWarningAsync(string title, string message);
        Task ShowInfoAsync(string title, string message);

        // Queue-specific UI updates
        Task UpdateQueueCountAsync(int count);
        Task RefreshQueueUIAsync();
    }

    // Context for UI state updates
    public class UIStateContext
    {
        public bool SaveQuotesMode { get; set; }
        public int PendingFileCount { get; set; }
        public bool AllFilesValid { get; set; }
        public bool IsProcessing { get; set; }
        public string? SelectedScope { get; set; }
        public bool HasCompanyName { get; set; }
        public string? CompanyNameInput { get; set; }
        public string? DetectedCompanyName { get; set; }
    }

    // Progress reporting for batch operations
    public class UIProgressUpdate
    {
        public double ProgressValue { get; set; }
        public string StatusMessage { get; set; } = "";
        public string? CurrentItem { get; set; }
        public int CompletedItems { get; set; }
        public int TotalItems { get; set; }
        public bool IsComplete { get; set; }
    }
} // End namespace DocHandler.Services 