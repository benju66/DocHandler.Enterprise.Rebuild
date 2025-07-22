// Folder: Services/
// File: OptimizedFileProcessingService.cs
// Enhanced file processing service using optimized Office conversion with pooling
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class OptimizedFileProcessingService : IOptimizedFileProcessingService
    {
        private readonly ILogger _logger = Log.ForContext<OptimizedFileProcessingService>();
        // COORDINATED SERVICES: Use shared session services instead of creating own instances
        private readonly SessionAwareOfficeService? _sharedWordService;
        private readonly SessionAwareExcelService? _sharedExcelService;
        private readonly PdfOperationsService _pdfOperationsService;
        private readonly ConfigurationService? _configService;
        private readonly PdfCacheService? _pdfCacheService;
        private readonly ProcessManager? _processManager;
        
        // Progress reporting delegate is now defined in IServices.cs
        
        // Add private disposal tracking
        private bool _disposed = false;

        private readonly HashSet<string> _supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".pdf", ".doc", ".docx", ".xls", ".xlsx"
        };

                public OptimizedFileProcessingService(
            ConfigurationService? configService = null, 
            PdfCacheService? pdfCacheService = null,
            ProcessManager? processManager = null,
            object? officeTracker = null, // Keep parameter for compatibility, but unused
            SessionAwareOfficeService? sharedWordService = null,
            SessionAwareExcelService? sharedExcelService = null)
        {
            // CRITICAL MEMORY FIX: Don't use shared session services - create converters on demand
            _sharedWordService = null; // Always null to force on-demand creation
            _sharedExcelService = null; // Always null to force on-demand creation
            _pdfOperationsService = new PdfOperationsService();
            _configService = configService;
            _pdfCacheService = pdfCacheService;
            _processManager = processManager;
            // officeTracker parameter kept for compatibility but no longer used
            
            _logger.Information("OptimizedFileProcessingService initialized with on-demand Office instance creation");
        }

        public bool IsFileSupported(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            return _supportedExtensions.Contains(extension);
        }

        public List<string> ValidateDroppedFiles(string[] files)
        {
            var validFiles = new List<string>();
            
            foreach (var file in files)
            {
                try
                {
                    // Use enhanced validation with custom exceptions
                    DocHandler.Helpers.FileValidator.ValidateFileOrThrow(file);
                    
                    if (IsFileSupported(file))
                    {
                        validFiles.Add(file);
                        _logger.Information("Valid file added: {FilePath}", file);
                    }
                    else
                    {
                        _logger.Warning("Unsupported file type: {FilePath}", file);
                    }
                }
                catch (SecurityViolationException secEx)
                {
                    _logger.Fatal(secEx, "Security violation detected in dropped file: {FilePath}", file);
                    // Security violations are never added to valid files
                }
                catch (FileValidationException fileEx)
                {
                    _logger.Warning(fileEx, "File validation failed for dropped file: {FilePath}", file);
                    // Validation failures are never added to valid files
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Unexpected error validating dropped file: {FilePath}", file);
                    // Unknown errors are never added to valid files
                }
            }
            
            return validFiles;
        }

        public async Task<ProcessingResult> ProcessFiles(List<string> filePaths, string outputDirectory, bool convertOfficeToPdf = true)
        {
            var result = new ProcessingResult();
            _logger.Information("Processing {Count} files to {OutputDirectory} with optimized conversion", filePaths.Count, outputDirectory);

            try
            {
                // Ensure output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Group files by type
                var fileGroups = filePaths.GroupBy(f => Path.GetExtension(f).ToLowerInvariant()).ToList();
                
                // Process Office files first if conversion is enabled
                var processedPdfs = new List<string>();
                var officeFiles = new List<string>();

                foreach (var group in fileGroups)
                {
                    switch (group.Key)
                    {
                        case ".doc":
                        case ".docx":
                            if (convertOfficeToPdf)
                            {
                                // Use optimized parallel processing for Word files
                                var wordResults = await ProcessWordFilesInParallelOptimized(group.ToList(), outputDirectory);
                                processedPdfs.AddRange(wordResults.SuccessfulFiles);
                                result.SuccessfulFiles.AddRange(wordResults.SuccessfulFiles);
                                result.FailedFiles.AddRange(wordResults.FailedFiles);
                                if (!string.IsNullOrEmpty(wordResults.ErrorMessage))
                                {
                                    result.ErrorMessage = wordResults.ErrorMessage;
                                }
                            }
                            else
                            {
                                officeFiles.AddRange(group);
                            }
                            break;

                        case ".xls":
                        case ".xlsx":
                            if (convertOfficeToPdf)
                            {
                                // Note: Excel conversion can be added later using similar optimized approach
                                // For now, use existing Excel conversion
                                var excelResults = await ProcessExcelFiles(group.ToList(), outputDirectory);
                                processedPdfs.AddRange(excelResults.SuccessfulFiles);
                                result.SuccessfulFiles.AddRange(excelResults.SuccessfulFiles);
                                result.FailedFiles.AddRange(excelResults.FailedFiles);
                                if (!string.IsNullOrEmpty(excelResults.ErrorMessage))
                                {
                                    result.ErrorMessage = excelResults.ErrorMessage;
                                }
                            }
                            else
                            {
                                officeFiles.AddRange(group);
                            }
                            break;

                        case ".pdf":
                            processedPdfs.AddRange(group);
                            break;
                    }
                }

                // If multiple PDFs, merge them
                if (processedPdfs.Count > 1)
                {
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var mergedFileName = $"Merged_Document_{timestamp}.pdf";
                    var mergedPath = Path.Combine(outputDirectory, mergedFileName);

                    _logger.Information("Merging {Count} PDFs into single document", processedPdfs.Count);
                    
                    if (await _pdfOperationsService.MergePdfFiles(processedPdfs, mergedPath))
                    {
                        result.SuccessfulFiles.Clear(); // Clear individual files
                        result.SuccessfulFiles.Add(mergedPath);
                        result.IsMerged = true;
                        _logger.Information("Successfully created merged PDF: {File}", mergedFileName);

                        // Clean up temporary converted PDFs (only those we created, not original PDFs)
                        foreach (var pdf in processedPdfs)
                        {
                            if (!filePaths.Contains(pdf)) // Don't delete original PDFs
                            {
                                try
                                {
                                    File.Delete(pdf);
                                    _logger.Debug("Cleaned up temporary file: {File}", Path.GetFileName(pdf));
                                }
                                catch (Exception ex)
                                {
                                    _logger.Warning(ex, "Failed to clean up temporary file: {File}", pdf);
                                }
                            }
                        }
                    }
                    else
                    {
                        result.ErrorMessage = "Failed to merge PDF files";
                        _logger.Error("PDF merge operation failed");
                    }
                }
                else if (processedPdfs.Count == 1)
                {
                    // Single PDF - just copy it if it's not already in the output directory
                    var sourcePdf = processedPdfs.First();
                    if (Path.GetDirectoryName(sourcePdf) != outputDirectory)
                    {
                        var outputPath = Path.Combine(outputDirectory, Path.GetFileName(sourcePdf));
                        File.Copy(sourcePdf, outputPath, true);
                        result.SuccessfulFiles.Add(outputPath);
                        _logger.Information("Copied single PDF to output: {File}", Path.GetFileName(outputPath));
                    }
                    else
                    {
                        result.SuccessfulFiles.Add(sourcePdf);
                    }
                }

                // Copy non-PDF office files if not converting
                foreach (var file in officeFiles)
                {
                    var outputPath = Path.Combine(outputDirectory, Path.GetFileName(file));
                    File.Copy(file, outputPath, true);
                    result.SuccessfulFiles.Add(outputPath);
                    _logger.Information("Copied file to output: {File}", Path.GetFileName(outputPath));
                }

                result.Success = result.SuccessfulFiles.Count > 0;
                result.OutputDirectory = outputDirectory;

                _logger.Information("Optimized processing complete. Processed: {Processed}, Failed: {Failed}", 
                    result.SuccessfulFiles.Count, result.FailedFiles.Count);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during optimized file processing");
                result.Success = false;
                result.ErrorMessage = $"Processing error: {ex.Message}";
                return result;
            }
        }

        private async Task<ProcessingResult> ProcessWordFilesInParallelOptimized(List<string> wordFiles, string outputDirectory)
        {
            var result = new ProcessingResult();
            
            // Use higher concurrency since we fixed the bottleneck
            var maxConcurrency = Math.Min(Environment.ProcessorCount, wordFiles.Count);
            _logger.Information("Processing {Count} Word files with {Concurrency} concurrent operations using ReliableOfficeConverter", 
                wordFiles.Count, maxConcurrency);

            var semaphore = new SemaphoreSlim(maxConcurrency);
            var progress = 0;
            var overallStopwatch = Stopwatch.StartNew();
            
            var tasks = wordFiles.Select(async file =>
            {
                await semaphore.WaitAsync();
                try
                {
                    var outputPath = Path.Combine(outputDirectory, 
                        Path.GetFileNameWithoutExtension(file) + ".pdf");
                    
                    var fileStopwatch = Stopwatch.StartNew();
                    
                    // CRITICAL MEMORY FIX: Use ReliableOfficeConverter for each task
                    ConversionResult conversionResult;
                    using (var converter = new ReliableOfficeConverter())
                    {
                        conversionResult = converter.ConvertWordToPdf(file, outputPath, singleUse: true);
                    } // Converter is disposed after each file
                    
                    fileStopwatch.Stop();
                    
                    var currentProgress = Interlocked.Increment(ref progress);
                    
                    lock (result)
                    {
                        if (conversionResult.Success)
                        {
                            result.SuccessfulFiles.Add(outputPath);
                            _logger.Information("Converted {File} ({Progress}/{Total}) in {ElapsedMs}ms using ReliableOfficeConverter", 
                                Path.GetFileName(file), currentProgress, wordFiles.Count, fileStopwatch.ElapsedMilliseconds);
                        }
                        else
                        {
                            result.FailedFiles.Add((file, conversionResult.ErrorMessage ?? "Unknown error"));
                            _logger.Warning("Failed to convert {File}: {Error}", 
                                Path.GetFileName(file), conversionResult.ErrorMessage);
                        }
                    }
                }
                catch (Exception ex)
                {
                    lock (result)
                    {
                        result.FailedFiles.Add((file, ex.Message));
                        _logger.Error(ex, "Failed to convert Word file: {File}", file);
                    }
                }
                finally
                {
                    semaphore.Release();
                }
            });
            
            await Task.WhenAll(tasks);
            overallStopwatch.Stop();
            
            _logger.Information("Optimized Word processing complete in {TotalMs}ms. Success: {Success}, Failed: {Failed}", 
                overallStopwatch.ElapsedMilliseconds, result.SuccessfulFiles.Count, result.FailedFiles.Count);
            
            return result;
        }

        private async Task<ProcessingResult> ProcessExcelFiles(List<string> excelFiles, string outputDirectory)
        {
            var result = new ProcessingResult();
            
            // Use ReliableOfficeConverter for Excel files as well
            using (var converter = new ReliableOfficeConverter())
            {
                foreach (var file in excelFiles)
                {
                    try
                    {
                        var outputPath = Path.Combine(outputDirectory, 
                            Path.GetFileNameWithoutExtension(file) + ".pdf");
                        
                        var conversionResult = converter.ConvertExcelToPdf(file, outputPath, singleUse: false);
                        
                        if (conversionResult.Success)
                        {
                            result.SuccessfulFiles.Add(outputPath);
                            _logger.Information("Converted Excel to PDF: {File}", Path.GetFileName(file));
                        }
                        else
                        {
                            result.FailedFiles.Add((file, conversionResult.ErrorMessage ?? "Unknown error"));
                        }
                    }
                    catch (Exception ex)
                    {
                        result.FailedFiles.Add((file, ex.Message));
                        _logger.Error(ex, "Failed to convert Excel file: {File}", file);
                    }
                }
            } // Converter disposed after all Excel files processed
            
            return result;
        }

        public string GetFileTypeDescription(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return extension switch
            {
                ".pdf" => "PDF Document",
                ".doc" => "Word Document (Legacy)",
                ".docx" => "Word Document",
                ".xls" => "Excel Spreadsheet (Legacy)",
                ".xlsx" => "Excel Spreadsheet",
                _ => "Unknown File Type"
            };
        }

        public string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        public string GetUniqueFileName(string directory, string fileName)
        {
            var name = Path.GetFileNameWithoutExtension(fileName);
            var extension = Path.GetExtension(fileName);
            var uniqueFileName = fileName;
            var counter = 1;
            
            while (File.Exists(Path.Combine(directory, uniqueFileName)))
            {
                uniqueFileName = $"{name} ({counter}){extension}";
                counter++;
            }
            
            return uniqueFileName;
        }

        public string CreateTempFolder()
        {
            var tempPath = Path.Combine(Path.GetTempPath(), "DocHandler", Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempPath);
            return tempPath;
        }

        public string CreateOutputFolder(string basePath)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            var outputPath = Path.Combine(basePath, $"DocHandler_Output_{timestamp}");
            Directory.CreateDirectory(outputPath);
            return outputPath;
        }

        // Single file conversion for Save Quotes Mode
        public async Task<ConversionResult> ConvertSingleFile(string inputPath, string outputPath)
        {
            return await ConvertSingleFile(inputPath, outputPath, null);
        }

        // Synchronous single file conversion for STA thread pool usage
        public ConversionResult ConvertSingleFileSync(string inputPath, string outputPath)
        {
            var fileName = Path.GetFileName(inputPath);
            var extension = Path.GetExtension(inputPath).ToLowerInvariant();
            
            _logger.Information("=== CONVERT SINGLE FILE SYNC START ===");
            _logger.Information("CONVERT: Input: {InputPath}", inputPath);
            _logger.Information("CONVERT: Output: {OutputPath}", outputPath);
            _logger.Information("CONVERT: Extension: {Extension}", extension);
            
            try
            {
                ConversionResult result;
                
                if (extension == ".pdf")
                {
                    File.Copy(inputPath, outputPath, true);
                    result = new ConversionResult { Success = true, OutputPath = outputPath };
                }
                else if (extension == ".doc" || extension == ".docx")
                {
                    // CRITICAL MEMORY FIX: Use ReliableOfficeConverter instead of shared services
                    using (var converter = new ReliableOfficeConverter())
                    {
                        _logger.Information("CONVERT: Word document detected, calling ReliableOfficeConverter synchronously...");
                        result = converter.ConvertWordToPdf(inputPath, outputPath, singleUse: true);
                        _logger.Information("CONVERT: ReliableOfficeConverter returned - Success: {Success}, Error: {Error}", 
                            result.Success, result.ErrorMessage ?? "None");
                    }
                }
                else if (extension == ".xls" || extension == ".xlsx")
                {
                    // CRITICAL MEMORY FIX: Use ReliableOfficeConverter instead of shared services
                    using (var converter = new ReliableOfficeConverter())
                    {
                        result = converter.ConvertExcelToPdf(inputPath, outputPath, singleUse: true);
                    }
                }
                else
                {
                    result = new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = $"Unsupported file type: {extension}"
                    };
                }
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to convert file synchronously: {InputPath}", inputPath);
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Conversion failed: {ex.Message}"
                };
            }
        }
        
        // Update ConvertSingleFile to report progress
        public async Task<ConversionResult> ConvertSingleFile(
            string inputPath, 
            string outputPath,
            ProgressCallback? progressCallback = null)
        {
            var fileName = Path.GetFileName(inputPath);
            var extension = Path.GetExtension(inputPath).ToLowerInvariant();
            
            _logger.Information("=== CONVERT SINGLE FILE START ===");
            _logger.Information("CONVERT: Input: {InputPath}", inputPath);
            _logger.Information("CONVERT: Output: {OutputPath}", outputPath);
            _logger.Information("CONVERT: Extension: {Extension}", extension);
            
            try
            {
                progressCallback?.Invoke(fileName, 0, "Starting conversion...");
                
                // Check cache first if enabled
                if (_configService?.Config.EnablePdfCaching == true && _pdfCacheService != null && extension != ".pdf")
                {
                    var fileHash = await ComputeFileHash(inputPath);
                    var cachedPdf = await _pdfCacheService.GetCachedPdfAsync(inputPath, fileHash);
                    
                    if (cachedPdf != null)
                    {
                        progressCallback?.Invoke(fileName, 50, "Using cached PDF...");
                        
                        // Copy cached PDF to output
                        await Task.Run(() => File.Copy(cachedPdf, outputPath, true));
                        
                        progressCallback?.Invoke(fileName, 100, "Completed (cached)");
                        
                        return new ConversionResult 
                        { 
                            Success = true, 
                            OutputPath = outputPath 
                        };
                    }
                }
                
                progressCallback?.Invoke(fileName, 20, "Converting to PDF...");
                
                ConversionResult result;
                
                if (extension == ".pdf")
                {
                    progressCallback?.Invoke(fileName, 50, "Copying PDF...");
                    File.Copy(inputPath, outputPath, true);
                    result = new ConversionResult { Success = true, OutputPath = outputPath };
                }
                else if (extension == ".doc" || extension == ".docx")
                {
                    // CRITICAL MEMORY FIX: Use ReliableOfficeConverter instead of shared services
                    using (var converter = new ReliableOfficeConverter())
                    {
                        _logger.Information("CONVERT: Word document detected, calling ReliableOfficeConverter...");
                        result = converter.ConvertWordToPdf(inputPath, outputPath, singleUse: true);
                        _logger.Information("CONVERT: ReliableOfficeConverter returned - Success: {Success}, Error: {Error}", 
                            result.Success, result.ErrorMessage ?? "None");
                    }
                }
                else if (extension == ".xls" || extension == ".xlsx")
                {
                    // CRITICAL MEMORY FIX: Use ReliableOfficeConverter instead of shared services
                    using (var converter = new ReliableOfficeConverter())
                    {
                        result = converter.ConvertExcelToPdf(inputPath, outputPath, singleUse: true);
                    }
                }
                else
                {
                    result = new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = $"Unsupported file type: {extension}"
                    };
                }
                
                if (result.Success)
                {
                    progressCallback?.Invoke(fileName, 80, "Finalizing...");
                    
                    // Add to cache if enabled
                    if (_configService?.Config.EnablePdfCaching == true && _pdfCacheService != null && extension != ".pdf")
                    {
                        var fileHash = await ComputeFileHash(inputPath);
                        await _pdfCacheService.AddToCacheAsync(inputPath, outputPath, fileHash);
                    }
                    
                    progressCallback?.Invoke(fileName, 100, "Completed");
                }
                else
                {
                    progressCallback?.Invoke(fileName, 100, $"Failed: {result.ErrorMessage}");
                }
                
                return result;
            }
            catch (Exception ex)
            {
                progressCallback?.Invoke(fileName, 100, $"Error: {ex.Message}");
                throw;
            }
        }
        
        // Add file hash computation for cache key
        private async Task<string> ComputeFileHash(string filePath)
        {
            using var md5 = System.Security.Cryptography.MD5.Create();
            using var stream = File.OpenRead(filePath);
            
            var hash = await md5.ComputeHashAsync(stream);
            return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
        }

        // REMOVED: OnQueueProcessingCompleted method no longer needed since we don't use shared services
        // Each ReliableOfficeConverter instance is disposed after use, ensuring proper cleanup
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                // Note: Shared services are disposed by their owner (MainViewModel)
                // Don't dispose shared services here as they may be used by other components
                // Note: PdfOperationsService doesn't implement IDisposable
                _logger.Information("Optimized file processing service disposed (shared services not disposed)");
            }
            
            _disposed = true;
        }

        ~OptimizedFileProcessingService()
        {
            Dispose(false);
        }
    }
} 