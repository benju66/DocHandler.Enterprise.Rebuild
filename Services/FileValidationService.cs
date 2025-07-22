using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Handles file validation logic extracted from MainViewModel (Fixed Version)
    /// </summary>
    public class FileValidationService : IFileValidationService
    {
        private readonly ILogger _logger;
        private readonly IOptimizedFileProcessingService _fileProcessingService;
        private readonly IConfigurationService _configService;

        public FileValidationService(
            IOptimizedFileProcessingService fileProcessingService,
            IConfigurationService configService)
        {
            _logger = Log.ForContext<FileValidationService>();
            _fileProcessingService = fileProcessingService ?? throw new ArgumentNullException(nameof(fileProcessingService));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
        }

        public async Task<ValidationResult> ValidateAsync(IEnumerable<FileItem> files, CancellationToken cancellationToken = default)
        {
            var result = new ValidationResult { IsValid = true };
            var fileList = files?.ToList() ?? new List<FileItem>();

            if (!fileList.Any())
            {
                result.IsValid = false;
                result.ErrorMessage = "No files provided for validation";
                return result;
            }

            _logger.Information("Starting validation of {FileCount} files", fileList.Count);

            try
            {
                foreach (var file in fileList)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var fileValidation = await ValidateSingleFileAsync(file, cancellationToken);
                    
                    if (fileValidation.IsValid)
                    {
                        result.ValidFiles.Add(file);
                        file.ValidationStatus = ValidationStatus.Valid;
                    }
                    else
                    {
                        result.InvalidFiles.Add(file);
                        if (!string.IsNullOrEmpty(fileValidation.ErrorMessage))
                        {
                            result.ErrorMessage = result.ErrorMessage == null 
                                ? fileValidation.ErrorMessage 
                                : $"{result.ErrorMessage}; {fileValidation.ErrorMessage}";
                        }
                        foreach (var warning in fileValidation.Warnings)
                        {
                            result.Warnings.Add(warning);
                        }
                        file.ValidationStatus = ValidationStatus.Invalid;
                        file.ValidationError = fileValidation.ErrorMessage;
                    }
                }

                result.IsValid = result.ValidFiles.Any() && result.InvalidFiles.Count == 0;

                _logger.Information("Validation completed: {ValidCount} valid, {InvalidCount} invalid files", 
                    result.ValidFiles.Count, result.InvalidFiles.Count);

                return result;
            }
            catch (OperationCanceledException)
            {
                _logger.Information("File validation was cancelled");
                result.IsValid = false;
                result.ErrorMessage = "Validation was cancelled";
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "File validation failed");
                result.IsValid = false;
                result.ErrorMessage = $"Validation failed: {ex.Message}";
                return result;
            }
        }

        public async Task<List<FileItem>> ValidateDroppedFilesAsync(string[] filePaths, CancellationToken cancellationToken = default)
        {
            var fileItems = new List<FileItem>();

            if (filePaths == null || !filePaths.Any())
            {
                _logger.Warning("No file paths provided for validation");
                return fileItems;
            }

            _logger.Information("Validating {FileCount} dropped files", filePaths.Length);

            try
            {
                // First, filter to supported files only
                var supportedFiles = _fileProcessingService.ValidateDroppedFiles(filePaths);
                
                foreach (var filePath in supportedFiles)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    try
                    {
                        if (!File.Exists(filePath))
                        {
                            _logger.Warning("File does not exist: {FilePath}", filePath);
                            continue;
                        }

                        var fileInfo = new FileInfo(filePath);
                        var fileItem = new FileItem
                        {
                            FilePath = filePath,
                            FileName = Path.GetFileName(filePath),
                            FileSize = fileInfo.Length,
                            FileType = Path.GetExtension(filePath).ToUpperInvariant().TrimStart('.'),
                            ValidationStatus = ValidationStatus.Pending
                        };

                        // Perform additional validation
                        var validation = await ValidateSingleFileAsync(fileItem, cancellationToken);
                        
                        if (validation.IsValid)
                        {
                            fileItem.ValidationStatus = ValidationStatus.Valid;
                            fileItems.Add(fileItem);
                            _logger.Debug("Valid file added: {FileName}", fileItem.FileName);
                        }
                        else
                        {
                            fileItem.ValidationStatus = ValidationStatus.Invalid;
                            fileItem.ValidationError = validation.ErrorMessage;
                            _logger.Warning("Invalid file skipped: {FileName} - {Error}", 
                                fileItem.FileName, validation.ErrorMessage);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to process dropped file: {FilePath}", filePath);
                    }
                }

                _logger.Information("Dropped file validation completed: {ValidCount} valid files", fileItems.Count);
                return fileItems;
            }
            catch (OperationCanceledException)
            {
                _logger.Information("Dropped file validation was cancelled");
                return fileItems;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Dropped file validation failed");
                return fileItems;
            }
        }

        public ValidationResult ValidateQuick(IEnumerable<FileItem> files)
        {
            var result = new ValidationResult { IsValid = true };
            var fileList = files?.ToList() ?? new List<FileItem>();

            if (!fileList.Any())
            {
                result.IsValid = false;
                result.ErrorMessage = "No files provided for validation";
                return result;
            }

            try
            {
                foreach (var file in fileList)
                {
                    // Quick validation without I/O operations
                    var quickValidation = ValidateQuickSingle(file);
                    
                    if (quickValidation.IsValid)
                    {
                        result.ValidFiles.Add(file);
                    }
                    else
                    {
                        result.InvalidFiles.Add(file);
                        if (!string.IsNullOrEmpty(quickValidation.ErrorMessage))
                        {
                            result.ErrorMessage = result.ErrorMessage == null 
                                ? quickValidation.ErrorMessage 
                                : $"{result.ErrorMessage}; {quickValidation.ErrorMessage}";
                        }
                        foreach (var warning in quickValidation.Warnings)
                        {
                            result.Warnings.Add(warning);
                        }
                    }
                }

                result.IsValid = result.ValidFiles.Any() && result.InvalidFiles.Count == 0;
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Quick validation failed");
                result.IsValid = false;
                result.ErrorMessage = $"Quick validation failed: {ex.Message}";
                return result;
            }
        }

        private async Task<ValidationResult> ValidateSingleFileAsync(FileItem file, CancellationToken cancellationToken = default)
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                // Check if file exists
                if (!File.Exists(file.FilePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File does not exist: {file.FilePath}";
                    return result;
                }

                // Check if file is supported
                if (!_fileProcessingService.IsFileSupported(file.FilePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File type not supported: {file.FileType}";
                    return result;
                }

                // Check file size - use the correct property name
                var fileInfo = new FileInfo(file.FilePath);
                var maxFileSize = 100 * 1024 * 1024; // Default 100MB
                
                // Try to get from config if available
                if (_configService.Config.DocFileSizeLimitMB > 0)
                {
                    maxFileSize = _configService.Config.DocFileSizeLimitMB * 1024 * 1024;
                }
                
                if (fileInfo.Length > maxFileSize)
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File size ({fileInfo.Length / (1024 * 1024)} MB) exceeds maximum allowed size ({maxFileSize / (1024 * 1024)} MB)";
                    return result;
                }

                // Check if file is locked
                if (await IsFileLockedAsync(file.FilePath, cancellationToken))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File is currently in use by another process: {file.FileName}";
                    return result;
                }

                // Additional file-specific validation can be added here
                
                result.IsValid = true;
                return result;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to validate file: {FilePath}", file.FilePath);
                result.IsValid = false;
                result.ErrorMessage = $"Validation error: {ex.Message}";
                return result;
            }
        }

        private ValidationResult ValidateQuickSingle(FileItem file)
        {
            var result = new ValidationResult { IsValid = true };

            try
            {
                // Basic validation without I/O
                if (string.IsNullOrWhiteSpace(file.FilePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "File path is empty";
                    return result;
                }

                if (string.IsNullOrWhiteSpace(file.FileName))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "File name is empty";
                    return result;
                }

                // Check if file type is supported (basic check)
                if (!_fileProcessingService.IsFileSupported(file.FilePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File type not supported: {file.FileType}";
                    return result;
                }

                result.IsValid = true;
                return result;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Quick validation failed for file: {FileName}", file.FileName);
                result.IsValid = false;
                result.ErrorMessage = $"Quick validation error: {ex.Message}";
                return result;
            }
        }

        private async Task<bool> IsFileLockedAsync(string filePath, CancellationToken cancellationToken = default)
        {
            try
            {
                await Task.Run(() =>
                {
                    using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        // If we can open the file, it's not locked
                    }
                }, cancellationToken);
                
                return false;
            }
            catch (IOException)
            {
                // File is locked
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                // File is locked or access denied
                return true;
            }
            catch (Exception)
            {
                // Assume file is accessible for other exceptions
                return false;
            }
        }
    }
} 