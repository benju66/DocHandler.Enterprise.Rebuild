using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services
{
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

    // Company Name Service Interface
    public interface ICompanyNameService : IDisposable
    {
        Task<string?> DetectCompanyNameAsync(string filePath);
        Task AddCompanyNameAsync(string companyName);
        Task RemoveCompanyNameAsync(string companyName);
        List<string> GetAllCompanyNames();
        void SetOfficeServices(ISessionAwareOfficeService officeService, ISessionAwareExcelService excelService);
    }

    // Session Aware Office Service Interface
    public interface ISessionAwareOfficeService : IDisposable
    {
        Task<string?> ConvertToPdfAsync(string inputPath, string outputPath);
        Task<string?> ExtractTextAsync(string filePath);
        Task<bool> IsAvailableAsync();
    }

    // Session Aware Excel Service Interface
    public interface ISessionAwareExcelService : IDisposable
    {
        Task<string?> ConvertToPdfAsync(string inputPath, string outputPath);
        Task<string?> ExtractTextAsync(string filePath);
        Task<bool> IsAvailableAsync();
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
        Task AddScopeAsync(ScopeOfWork scope);
        Task RemoveScopeAsync(int id);
        Task UpdateScopeAsync(ScopeOfWork scope);
        Task AddToRecentAsync(string scopeName);
    }

    // Process Manager Interface
    public interface IProcessManager
    {
        Task<bool> KillOfficeProcessesAsync();
        Task<List<System.Diagnostics.Process>> GetOfficeProcessesAsync();
        Task<bool> IsOfficeRunningAsync();
        void RegisterOfficeProcess(int processId);
        void UnregisterOfficeProcess(int processId);
        Task CleanupOrphanedProcessesAsync();
    }

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
        void AddToQueue(FileItem file, string scope, string companyName, string saveLocation);
        Task StartProcessingAsync();
        Task StopProcessingAsync();
        Task ClearQueueAsync();
        bool IsProcessing { get; }
        int TotalCount { get; }
        int ProcessedCount { get; }
        int FailedCount { get; }
        
        event EventHandler<SaveQuoteProgressEventArgs> ProgressChanged;
        event EventHandler<SaveQuoteCompletedEventArgs> ItemCompleted;
        event EventHandler QueueEmpty;
        event EventHandler<string> StatusMessageChanged;
    }

    // Health Monitoring Interfaces
    public interface IOfficeHealthMonitor : IDisposable
    {
        Task<bool> CheckOfficeHealthAsync();
        void StartMonitoring();
        void StopMonitoring();
    }

    public interface IApplicationHealthChecker
    {
        Task<HealthCheckResult> CheckSystemHealthAsync();
        Task<bool> CheckOfficeInstallationAsync();
        Task<bool> CheckDependenciesAsync();
    }

    // Circuit Breaker Interfaces
    public interface ICircuitBreaker
    {
        Task<T> ExecuteAsync<T>(Func<Task<T>> operation);
        void Reset();
        bool IsOpen { get; }
    }

    public interface IConversionCircuitBreaker : ICircuitBreaker
    {
        Task<string?> ExecuteConversionAsync(Func<Task<string?>> conversion);
    }
}

// Supporting classes and events are defined in their respective service files 