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
    /// Handles company name detection logic extracted from MainViewModel
    /// </summary>
    public class CompanyDetectionService : IEnhancedCompanyDetectionService
    {
        private readonly ILogger _logger;
        private readonly ICompanyNameService _companyNameService;
        private readonly IConfigurationService _configService;
        private readonly SemaphoreSlim _scanSemaphore;
        private volatile int _activeScanCount = 0;

        public CompanyDetectionService(
            ICompanyNameService companyNameService,
            IConfigurationService configService)
        {
            _logger = Log.ForContext<CompanyDetectionService>();
            _companyNameService = companyNameService ?? throw new ArgumentNullException(nameof(companyNameService));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _scanSemaphore = new SemaphoreSlim(1, 1);
            
            _logger.Debug("CompanyDetectionService initialized");
        }

        public async Task<string?> DetectCompanyAsync(CompanyDetectionRequest request)
        {
            // Pipeline method - delegate to existing scan method
            return await ScanForCompanyNameAsync(request);
        }

        public async Task<string?> ScanForCompanyNameAsync(CompanyDetectionRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrWhiteSpace(request.FilePath))
            {
                _logger.Warning("Cannot scan for company name: file path is empty");
                return null;
            }

            // Try to acquire semaphore with timeout to prevent deadlocks
            if (!await _scanSemaphore.WaitAsync(TimeSpan.FromSeconds(5)))
            {
                _logger.Warning("Timeout waiting for scan semaphore: {FilePath}", request.FilePath);
                return null;
            }
            
            try
            {
                // Check concurrent scan limit inside the lock to prevent race conditions
                if (_activeScanCount >= 3)
                {
                    _logger.Debug("Too many concurrent scans, skipping: {FilePath}", request.FilePath);
                    return null;
                }
                
                Interlocked.Increment(ref _activeScanCount);
                
                _logger.Debug("Starting company name scan for: {FilePath}", request.FilePath);

                var companyName = await _companyNameService.ScanDocumentForCompanyName(
                    request.FilePath, 
                    null); // No progress reporting for legacy method

                if (!string.IsNullOrWhiteSpace(companyName))
                {
                    _logger.Information("Company name detected: {CompanyName} in {FilePath}", 
                        companyName, request.FilePath);
                }
                else
                {
                    _logger.Debug("No company name detected in: {FilePath}", request.FilePath);
                }

                return companyName;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error scanning for company name: {FilePath}", request.FilePath);
                return null;
            }
            finally
            {
                Interlocked.Decrement(ref _activeScanCount);
                _scanSemaphore.Release();
            }
        }

        public async Task<Dictionary<string, string?>> ScanMultipleFilesAsync(
            IEnumerable<string> filePaths, 
            IProgress<int>? progress = null, 
            CancellationToken cancellationToken = default)
        {
            var results = new Dictionary<string, string?>();
            var fileList = filePaths?.ToList() ?? new List<string>();

            if (!fileList.Any())
            {
                _logger.Warning("No files provided for company name scanning");
                return results;
            }

            _logger.Information("Starting company name scan for {FileCount} files", fileList.Count);

            try
            {
                var totalFiles = fileList.Count;
                var processedFiles = 0;

                foreach (var filePath in fileList)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        var request = new CompanyDetectionRequest
                        {
                            FilePath = filePath,
                            UseCache = true,
                            TimeoutSeconds = 30
                        };
                        var companyName = await ScanForCompanyNameAsync(request);
                        results[filePath] = companyName;

                        processedFiles++;
                        var progressPercentage = (int)((double)processedFiles / totalFiles * 100);
                        progress?.Report(progressPercentage);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to scan file for company name: {FilePath}", filePath);
                        results[filePath] = null;
                    }
                }

                _logger.Information("Company name scanning completed: {ResultCount} results", results.Count);
                return results;
            }
            catch (OperationCanceledException)
            {
                _logger.Information("Company name scanning was cancelled");
                return results;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Company name scanning failed");
                return results;
            }
        }

        public bool ShouldAutoScan(FileItem file)
        {
            if (file == null)
                return false;

            // Check if auto-scanning is enabled in configuration
            if (!_configService.Config.AutoScanCompanyNames)
            {
                return false;
            }

            // Check if we're in Save Quotes mode
            if (!_configService.Config.SaveQuotesMode)
            {
                return false;
            }

            // Check file type - only scan supported document types
            var supportedExtensions = new[] { ".pdf", ".doc", ".docx" };
            var extension = System.IO.Path.GetExtension(file.FilePath)?.ToLowerInvariant();
            
            if (!supportedExtensions.Contains(extension))
            {
                return false;
            }

            // Check file size for .doc files specifically
            if (extension == ".doc")
            {
                if (!_configService.Config.ScanCompanyNamesForDocFiles)
                {
                    return false;
                }

                try
                {
                    var fileInfo = new System.IO.FileInfo(file.FilePath);
                    var fileSizeMB = fileInfo.Length / (1024.0 * 1024.0);
                    
                    if (fileSizeMB > _configService.Config.DocFileSizeLimitMB)
                    {
                        _logger.Debug("Skipping .doc file scan due to size limit: {FileSizeMB}MB > {LimitMB}MB", 
                            fileSizeMB, _configService.Config.DocFileSizeLimitMB);
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to check file size for auto-scan decision: {FilePath}", file.FilePath);
                    return false;
                }
            }

            return true;
        }

        public void ClearDetectionCache()
        {
            try
            {
                _logger.Information("Clearing company name detection cache");
                _companyNameService.CleanupPdfCache();
                _logger.Information("Company name detection cache cleared");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to clear company name detection cache");
            }
        }

        // Enhanced interface implementations
        public async Task<EnhancedCompanyDetectionResult> DetectCompanyAsync(string filePath, IProgress<int>? progress = null, CancellationToken cancellationToken = default)
        {
            var startTime = DateTime.UtcNow;
            try
            {
                var request = new CompanyDetectionRequest
                {
                    FilePath = filePath,
                    UseCache = true,
                    TimeoutSeconds = 30
                };

                var detectedCompany = await ScanForCompanyNameAsync(request);
                
                return new EnhancedCompanyDetectionResult
                {
                    FilePath = filePath,
                    DetectedCompany = detectedCompany,
                    Confidence = !string.IsNullOrEmpty(detectedCompany) ? 0.85 : 0.0,
                    ProcessingTime = DateTime.UtcNow - startTime
                };
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error detecting company for file: {FilePath}", filePath);
                return new EnhancedCompanyDetectionResult
                {
                    FilePath = filePath,
                    ErrorMessage = ex.Message,
                    ProcessingTime = DateTime.UtcNow - startTime
                };
            }
        }

        public async Task<List<EnhancedCompanyDetectionResult>> DetectCompaniesAsync(IEnumerable<string> filePaths, IProgress<BatchProgress>? progress = null, CancellationToken cancellationToken = default)
        {
            var results = new List<EnhancedCompanyDetectionResult>();
            var filePathList = filePaths.ToList();
            var completed = 0;

            foreach (var filePath in filePathList)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var result = await DetectCompanyAsync(filePath, null, cancellationToken);
                results.Add(result);
                
                completed++;
                progress?.Report(new BatchProgress
                {
                    TotalItems = filePathList.Count,
                    CompletedItems = completed,
                    CurrentItem = System.IO.Path.GetFileName(filePath)
                });
            }

            return results;
        }

        public async Task<CompanyValidationResult> ValidateCompanyAsync(string companyName)
        {
            await Task.Yield(); // Make it async
            
            // Simple validation - in a real implementation this would check against a database
            var isValid = !string.IsNullOrWhiteSpace(companyName) && companyName.Length >= 2;
            
            return new CompanyValidationResult
            {
                CompanyName = companyName,
                IsKnownCompany = isValid,
                IsValid = isValid,
                UsageCount = isValid ? 1 : 0,
                LastUsed = DateTime.UtcNow
            };
        }

        public async Task<List<string>> GetCompanySuggestionsAsync(string partialName, int maxSuggestions = 10)
        {
            await Task.Yield(); // Make it async
            
            // Simple implementation - in a real system this would use fuzzy matching
            var suggestions = new List<string>();
            
            if (!string.IsNullOrWhiteSpace(partialName))
            {
                // Add some mock suggestions based on partial name
                suggestions.Add($"{partialName} Corp");
                suggestions.Add($"{partialName} Inc");
                suggestions.Add($"{partialName} LLC");
            }
            
            return suggestions.Take(maxSuggestions).ToList();
        }

        public async Task<bool> AddCompanyAsync(string companyName, List<string>? aliases = null)
        {
            await Task.Yield(); // Make it async
            
            if (string.IsNullOrWhiteSpace(companyName))
                return false;
                
            _logger.Information("Would add company: {CompanyName} with {AliasCount} aliases", 
                companyName, aliases?.Count ?? 0);
            
            return true; // Simulate success
        }

        public async Task IncrementCompanyUsageAsync(string companyName)
        {
            await Task.Yield(); // Make it async
            
            if (!string.IsNullOrWhiteSpace(companyName))
            {
                _logger.Debug("Would increment usage for company: {CompanyName}", companyName);
            }
        }
    }
} 