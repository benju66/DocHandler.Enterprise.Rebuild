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
    public class CompanyDetectionService : ICompanyDetectionService
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

            if (!ShouldAutoScan(new FileItem { FilePath = request.FilePath }))
            {
                _logger.Debug("Auto-scan disabled or not applicable for file: {FilePath}", request.FilePath);
                return null;
            }

            // Protect against too many concurrent scans
            if (_activeScanCount >= 3)
            {
                _logger.Debug("Too many concurrent scans, skipping: {FilePath}", request.FilePath);
                return null;
            }

            await _scanSemaphore.WaitAsync(request.CancellationToken);
            
            try
            {
                Interlocked.Increment(ref _activeScanCount);
                
                _logger.Debug("Starting company name scan for: {FilePath}", request.FilePath);
                request.Progress?.Report(10);

                var companyName = await _companyNameService.ScanDocumentForCompanyName(
                    request.FilePath, 
                    request.Progress);

                request.Progress?.Report(100);

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
            catch (OperationCanceledException)
            {
                _logger.Debug("Company name scan cancelled for: {FilePath}", request.FilePath);
                return null;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Company name scan failed for: {FilePath}", request.FilePath);
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
                        var request = new CompanyDetectionRequest(filePath, null, cancellationToken);
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
    }
} 