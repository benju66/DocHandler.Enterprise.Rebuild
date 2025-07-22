using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;
using DocHandler.Helpers;
using Serilog;

namespace DocHandler.Services.Modes
{
    /// <summary>
    /// Save Quotes processing mode - extracts quote documents to organized folders
    /// </summary>
    public class SaveQuotesMode : ProcessingModeBase
    {
        public override string ModeName => "SaveQuotes";
        public override string DisplayName => "Save Quotes";
        public override string Description => "Organize and save quote documents with company names and scope of work";
        public override Version Version => new Version(1, 0, 0);

        // Required services
        private SaveQuotesQueueService? _queueService;
        private CompanyNameService? _companyNameService;
        private ScopeOfWorkService? _scopeOfWorkService;
        private ConfigurationService? _configService;
        private OptimizedFileProcessingService? _fileProcessingService;
        private PdfCacheService? _pdfCacheService;
        private ProcessManager? _processManager;

        public SaveQuotesMode()
        {
            // Constructor - services will be injected during initialization
        }

        protected override async Task InitializeModeAsync()
        {
            // Get required services from DI container
            _queueService = GetService<SaveQuotesQueueService>();
            _companyNameService = GetService<CompanyNameService>();
            _scopeOfWorkService = GetService<ScopeOfWorkService>();
            _configService = GetService<ConfigurationService>();
            _fileProcessingService = GetService<OptimizedFileProcessingService>();
            _pdfCacheService = GetService<PdfCacheService>();
            _processManager = GetService<ProcessManager>();

            // Validate required services
            if (_queueService == null)
                throw new InvalidOperationException("SaveQuotesQueueService is required for SaveQuotes mode");
                
            if (_companyNameService == null)
                throw new InvalidOperationException("CompanyNameService is required for SaveQuotes mode");
            
            if (_scopeOfWorkService == null)
                throw new InvalidOperationException("ScopeOfWorkService is required for SaveQuotes mode");
            
            if (_configService == null)
                throw new InvalidOperationException("ConfigurationService is required for SaveQuotes mode");

            // Load required data
            await _companyNameService.LoadDataAsync();
            await _scopeOfWorkService.LoadDataAsync();

            _logger.Information("SaveQuotes mode initialized successfully");
        }

        protected override async Task<ModeProcessingResult> ProcessFilesAsync(ProcessingRequest request, CancellationToken cancellationToken)
        {
            var startTime = DateTime.UtcNow;
            var processedFiles = new List<ProcessedFile>();

            try
            {
                // Validate required parameters
                if (!request.Parameters.TryGetValue("scope", out var scopeObj) || scopeObj is not string scope || string.IsNullOrWhiteSpace(scope))
                {
                    return new ModeProcessingResult
                    {
                        Success = false,
                        ErrorMessage = "Scope of work parameter is required for SaveQuotes mode",
                        ProcessedFiles = processedFiles,
                        Duration = DateTime.UtcNow - startTime
                    };
                }

                if (!request.Parameters.TryGetValue("companyName", out var companyObj) || companyObj is not string companyName || string.IsNullOrWhiteSpace(companyName))
                {
                    return new ModeProcessingResult
                    {
                        Success = false,
                        ErrorMessage = "Company name parameter is required for SaveQuotes mode",
                        ProcessedFiles = processedFiles,
                        Duration = DateTime.UtcNow - startTime
                    };
                }

                // Sanitize company name for filename usage
                var sanitizedCompanyName = SanitizeFileName(companyName);

                // Get or create queue service
                if (_queueService == null)
            {
                throw new InvalidOperationException("SaveQuotesQueueService is not available");
            }
            
            var queueService = _queueService;

                // Add files to queue
                foreach (var file in request.Files)
                {
                    queueService.AddToQueue(file, scope, sanitizedCompanyName, request.OutputDirectory);
                    
                    processedFiles.Add(new ProcessedFile
                    {
                        OriginalFile = file,
                        Success = true,
                        Metadata = new Dictionary<string, object>
                        {
                            ["scope"] = scope,
                            ["companyName"] = sanitizedCompanyName,
                            ["status"] = "queued"
                        }
                    });
                }

                // Start processing if not already running
                if (!queueService.IsProcessing)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await queueService.StartProcessingAsync();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Failed to start queue processing in SaveQuotes mode");
                        }
                    }, cancellationToken);
                }

                _logger.Information("SaveQuotes mode processed {FileCount} files, added to queue", request.Files.Count);

                return new ModeProcessingResult
                {
                    Success = true,
                    ProcessedFiles = processedFiles,
                    Duration = DateTime.UtcNow - startTime,
                    Metadata = new Dictionary<string, object>
                    {
                        ["queuedFiles"] = request.Files.Count,
                        ["scope"] = scope,
                        ["companyName"] = sanitizedCompanyName
                    }
                };
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error processing files in SaveQuotes mode");
                
                return new ModeProcessingResult
                {
                    Success = false,
                    ErrorMessage = ex.Message,
                    Exception = ex,
                    ProcessedFiles = processedFiles,
                    Duration = DateTime.UtcNow - startTime
                };
            }
        }

        protected override bool IsFileSupported(FileItem file)
        {
            if (string.IsNullOrWhiteSpace(file.FilePath))
                return false;

            var extension = System.IO.Path.GetExtension(file.FilePath).ToLowerInvariant();
            
            // Support common document formats for quotes
            var supportedExtensions = new HashSet<string>
            {
                ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".txt", ".rtf",
                ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".tif"
            };

            return supportedExtensions.Contains(extension);
        }



        /// <summary>
        /// Sanitize a filename to remove invalid characters
        /// </summary>
        private static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "Unknown";

            var invalidChars = System.IO.Path.GetInvalidFileNameChars();
            var sanitized = fileName;
            
            foreach (var invalidChar in invalidChars)
            {
                sanitized = sanitized.Replace(invalidChar, '_');
            }
            
            // Also replace some additional problematic characters
            sanitized = sanitized.Replace(' ', '_')
                                 .Replace('.', '_')
                                 .Replace('-', '_');
            
            // Ensure it's not too long
            if (sanitized.Length > 50)
            {
                sanitized = sanitized.Substring(0, 50);
            }
            
            // Ensure it's not empty
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                sanitized = "Unknown";
            }
            
            return sanitized;
        }

        /// <summary>
        /// Get mode-specific configuration
        /// </summary>
        public override IModeConfiguration GetConfiguration()
        {
            var config = new ModeConfiguration(ModeName);
            
            // Add SaveQuotes-specific configuration options
            config.SetSetting("autoScanCompanyNames", _configService?.Config.AutoScanCompanyNames ?? true);
            config.SetSetting("clearScopeAfterProcessing", _configService?.Config.ClearScopeAfterProcessing ?? false);
            config.SetSetting("scanCompanyNamesForDocFiles", _configService?.Config.ScanCompanyNamesForDocFiles ?? false);
            config.SetSetting("docFileSizeLimitMB", _configService?.Config.DocFileSizeLimitMB ?? 10);
            
            return config;
        }

        /// <summary>
        /// Dispose of resources
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _queueService?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
} 