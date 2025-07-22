using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services
{
    /// <summary>
    /// Progress callback delegate for file processing operations
    /// </summary>
    public delegate void ProgressCallback(string fileName, double percentage, string status);
    // Configuration Service Interface
    public interface IConfigurationService
    {
        AppConfiguration Config { get; }
        Task SaveConfigurationAsync();
        void UpdateConfiguration(Action<AppConfiguration> updateAction);
    }

    // Office Conversion Service Interface
    public interface IOfficeConversionService : IDisposable
    {
        Task<string?> ConvertToPdfAsync(string inputPath, string outputPath);
        Task<bool> IsOfficeAvailableAsync();
        void Cleanup();
    }

    // File Processing Service Interface
    public interface IFileProcessingService : IDisposable
    {
        bool IsFileSupported(string filePath);
        Task<List<string>> ProcessFilesAsync(IEnumerable<string> filePaths, IProgress<string>? progress = null);
        Task<string?> ConvertToPdfAsync(string inputPath, string outputPath);
    }

    // Optimized File Processing Service Interface
    public interface IOptimizedFileProcessingService : IDisposable
    {
        bool IsFileSupported(string filePath);
        List<string> ValidateDroppedFiles(string[] files);
        Task<ProcessingResult> ProcessFiles(List<string> filePaths, string outputDirectory, bool convertOfficeToPdf = true);
        string GetFileTypeDescription(string filePath);
        string FormatFileSize(long bytes);
        string GetUniqueFileName(string directory, string fileName);
        string CreateTempFolder();
        string CreateOutputFolder(string basePath);
        Task<ConversionResult> ConvertSingleFile(string inputPath, string outputPath);
        ConversionResult ConvertSingleFileSync(string inputPath, string outputPath);
        Task<ConversionResult> ConvertSingleFile(string inputPath, string outputPath, ProgressCallback? progressCallback = null);
        Task<bool> ProcessFileAsync(string inputPath, string outputPath, ProgressCallback? progressCallback = null);
    }

    // Company Name Service Interface
    public interface ICompanyNameService : IDisposable
    {
        void SetOfficeServices(SessionAwareOfficeService officeService, SessionAwareExcelService excelService);
        void UpdateDocFileSizeLimit(int limitMB);
        Task LoadDataAsync();
        Task SaveCompanyNames();
        Task<bool> AddCompanyName(string name, List<string>? aliases = null);
        Task<bool> RemoveCompanyName(string name);
        Task<string?> ScanDocumentForCompanyName(string filePath, IProgress<int>? progress = null);
        Task IncrementUsageCount(string companyName);
        List<CompanyInfo> GetMostUsedCompanies(int count = 10);
        List<CompanyInfo> SearchCompanies(string searchTerm);
        string GetPerformanceSummary();
        bool TryGetCachedPdf(string originalFilePath, out string? cachedPdfPath);
        void CleanupPdfCache();
        string? GetCachedPdfPath(string originalFilePath);
        void RemoveCachedPdf(string originalFilePath);
    }

    // Session Aware Office Service Interface
    public interface ISessionAwareOfficeService : IDisposable
    {
        Task<ConversionResult> ConvertWordToPdf(string inputPath, string outputPath);
        void ForceCleanupIfIdle();
        bool IsOfficeInstalled();
    }

    // Session Aware Excel Service Interface
    public interface ISessionAwareExcelService : IDisposable
    {
        Task<ConversionResult> ConvertSpreadsheetToPdf(string inputPath, string outputPath);
        void ForceCleanupIfIdle();
        void DisposeIfIdle();
    }

    // PDF Operations Service Interface
    public interface IPdfOperationsService
    {
        Task<bool> MergePdfFiles(List<string> inputPaths, string outputPath);
        Task<string?> ExtractTextFromPdfAsync(string pdfPath);
        Task<bool> ValidatePdfAsync(string pdfPath);
    }

    // PDF Cache Service Interface
    public interface IPdfCacheService : IDisposable
    {
        Task<string?> GetCachedPdfAsync(string originalPath, string fileHash);
        Task CachePdfAsync(string originalPath, string pdfPath, string fileHash);
        Task ClearCacheAsync();
        Task<long> GetCacheSizeAsync();
    }

    // Scope of Work Service Interface
    public interface IScopeOfWorkService
    {
        List<ScopeOfWork> Scopes { get; }
        List<string> RecentScopes { get; }

        Task LoadDataAsync();
        Task SaveScopesOfWork();
        Task SaveRecentScopes();
        Task<bool> AddScope(string code, string description);
        Task<bool> UpdateScope(string oldCode, string newCode, string newDescription);
        Task<bool> RemoveScope(string code);
        Task UpdateRecentScope(string scopeText);
        Task ClearRecentScopes();
        Task IncrementUsageCount(string scopeText);
        string GetFormattedScope(ScopeOfWork scope);
        List<ScopeOfWork> SearchScopes(string searchTerm);
        List<ScopeOfWork> GetMostUsedScopes(int count = 10);
        Task<ImportResult> ImportScopes(string filePath, bool replace = false);
        Task<bool> ExportScopes(string filePath);
        Task<bool> ResetToDefaults();
    }

    // Process Manager Interface - extends existing interface

    // Performance Monitor Interface
    public interface IPerformanceMonitor : IDisposable
    {
        void StartOperation(string operationName);
        void EndOperation(string operationName);
        void RecordMetric(string metricName, double value);
        Task<Dictionary<string, object>> GetMetricsAsync();
        event EventHandler<MemoryPressureEventArgs>? MemoryPressureDetected;
    }

    // Telemetry Service Interface
    public interface ITelemetryService : IDisposable
    {
        void TrackEvent(string eventName, Dictionary<string, object>? properties = null);
        void TrackMetric(string metricName, double value, Dictionary<string, object>? properties = null);
        void TrackException(Exception exception, Dictionary<string, object>? properties = null);
        Task FlushAsync();
    }

    // Office Service Factory Interface
    public interface IOfficeServiceFactory
    {
        Task<IOfficeConversionService> CreateOfficeServiceAsync();
        Task<ISessionAwareOfficeService> CreateSessionOfficeServiceAsync();
        Task<ISessionAwareExcelService> CreateSessionExcelServiceAsync();
        Task<bool> IsOfficeAvailableAsync();
    }

    // Save Quotes Queue Service Interface
    public interface ISaveQuotesQueueService : IDisposable
    {
        System.Collections.ObjectModel.ObservableCollection<SaveQuoteItem> AllItems { get; }
        int TotalCount { get; set; }
        int ProcessedCount { get; set; }
        int FailedCount { get; set; }
        bool IsProcessing { get; set; }

        event EventHandler<SaveQuoteProgressEventArgs>? ProgressChanged;
        event EventHandler<SaveQuoteCompletedEventArgs>? ItemCompleted;
        event EventHandler? QueueEmpty;
        event EventHandler<string>? StatusMessageChanged;

        void StopProcessing();
        void AddToQueue(FileItem file, string scope, string companyName, string saveLocation);
        Task StartProcessingAsync();
        void UpdateMaxConcurrency(int newMax);
        void CancelItem(SaveQuoteItem item);
        void ClearCompleted();
    }

    // Mode Management Interfaces are defined in their respective files:
    // - IModeManager is defined in Services/ModeManager.cs
    // - IModeRegistry is defined in Services/ModeRegistry.cs
}

// Supporting classes and events are defined in their respective service files 