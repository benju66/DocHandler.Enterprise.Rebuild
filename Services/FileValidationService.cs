using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Serilog;
using DocHandler.Helpers;
using DocHandler.Services.Configuration;
using DocHandler.Models;

namespace DocHandler.Services
{
    /// <summary>
    /// Enhanced file validation service with comprehensive security and performance checks
    /// </summary>
    public class FileValidationService : IFileValidationService
    {
        private readonly ILogger _logger;
        private readonly IHierarchicalConfigurationService _configService;
        private readonly SemaphoreSlim _validationSemaphore;
        private readonly Dictionary<string, string> _mimeTypeMap;
        
        public FileValidationService(IHierarchicalConfigurationService configService)
        {
            _logger = Log.ForContext<FileValidationService>();
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _validationSemaphore = new SemaphoreSlim(Environment.ProcessorCount, Environment.ProcessorCount);
            _mimeTypeMap = InitializeMimeTypeMap();
            
            _logger.Information("FileValidationService initialized");
        }

        public async Task<EnhancedFileValidationResult> ValidateFileAsync(string filePath, CancellationToken cancellationToken = default)
        {
            var startTime = DateTime.UtcNow;
            var result = new EnhancedFileValidationResult
            {
                FilePath = filePath,
                ValidationTime = startTime
            };

            try
            {
                await _validationSemaphore.WaitAsync(cancellationToken);
                
                // Basic existence and access checks
                if (!File.Exists(filePath))
                {
                    result.IsValid = false;
                    result.Errors.Add("File does not exist");
                    return result;
                }

                var fileInfo = new FileInfo(filePath);
                result.FileSizeBytes = fileInfo.Length;
                result.FileType = Path.GetExtension(filePath).ToLowerInvariant();

                // Security validation
                var securityAssessment = await AssessSecurityRiskAsync(filePath);
                result.RiskLevel = securityAssessment.RiskLevel;
                result.IsSecure = securityAssessment.RiskLevel < SecurityRiskLevel.High;
                
                if (securityAssessment.IsBlocked)
                {
                    result.IsValid = false;
                    result.Errors.Add($"File blocked due to security risk: {securityAssessment.Recommendation}");
                    return result;
                }

                // File type validation
                if (!IsFileTypeSupported(filePath))
                {
                    result.IsValid = false;
                    result.Errors.Add($"Unsupported file type: {result.FileType}");
                    return result;
                }

                // Size validation - parse from string format
                var maxSizeStr = _configService.Config.ModeDefaults.MaxFileSize ?? "100MB";
                var maxSizeMB = ParseFileSizeToMB(maxSizeStr);
                var maxSizeBytes = maxSizeMB * 1024L * 1024L;
                if (fileInfo.Length > maxSizeBytes)
                {
                    result.IsValid = false;
                    result.Errors.Add($"File size ({FormatFileSize(fileInfo.Length)}) exceeds maximum allowed size ({maxSizeMB}MB)");
                    return result;
                }

                // Content validation (always enabled for security)
                await ValidateFileContentAsync(filePath, result, cancellationToken);

                // Path validation
                var sanitizedPath = SanitizeFilePath(filePath);
                if (sanitizedPath != filePath)
                {
                    result.Warnings.Add("File path contains potentially unsafe characters");
                }

                result.IsValid = result.Errors.Count == 0;
                
                _logger.Debug("File validation completed for {FilePath}: {IsValid} (Errors: {ErrorCount}, Warnings: {WarningCount})",
                    filePath, result.IsValid, result.Errors.Count, result.Warnings.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error validating file: {FilePath}", filePath);
                result.IsValid = false;
                result.Errors.Add($"Validation error: {ex.Message}");
            }
            finally
            {
                result.ValidationDuration = DateTime.UtcNow - startTime;
                _validationSemaphore.Release();
            }

            return result;
        }

        public async Task<List<EnhancedFileValidationResult>> ValidateFilesAsync(
            IEnumerable<string> filePaths, 
            IProgress<ValidationProgress>? progress = null, 
            CancellationToken cancellationToken = default)
        {
            var filePathList = filePaths.ToList();
            var results = new List<EnhancedFileValidationResult>();
            var completed = 0;

            var validationProgress = new ValidationProgress
            {
                TotalFiles = filePathList.Count
            };

            _logger.Information("Starting batch validation of {FileCount} files", filePathList.Count);

            var semaphore = new SemaphoreSlim(Math.Min(Environment.ProcessorCount, 10), Math.Min(Environment.ProcessorCount, 10));
            var tasks = filePathList.Select(async filePath =>
            {
                await semaphore.WaitAsync(cancellationToken);
                try
                {
                    var result = await ValidateFileAsync(filePath, cancellationToken);
                    
                    lock (results)
                    {
                        results.Add(result);
                        completed++;
                        
                        if (progress != null)
                        {
                            validationProgress.CompletedFiles = completed;
                            validationProgress.CurrentFile = Path.GetFileName(filePath);
                            progress.Report(validationProgress);
                        }
                    }
                    
                    return result;
                }
                finally
                {
                    semaphore.Release();
                }
            });

            await Task.WhenAll(tasks);
            
            var validFiles = results.Count(r => r.IsValid);
            _logger.Information("Batch validation completed: {ValidFiles}/{TotalFiles} files valid", 
                validFiles, filePathList.Count);

            return results.OrderBy(r => r.FilePath).ToList();
        }

        public bool IsFileTypeSupported(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            // Use a predefined list of supported extensions for now
            var allowedExtensions = new[] { ".pdf", ".docx", ".doc", ".xlsx", ".xls", ".txt", ".rtf" };
            
            return allowedExtensions.Contains(extension);
        }

        public async Task<SecurityRiskAssessment> AssessSecurityRiskAsync(string filePath)
        {
            var assessment = new SecurityRiskAssessment();
            var riskFactors = new List<string>();

            try
            {
                var fileName = Path.GetFileName(filePath);
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                var fileInfo = new FileInfo(filePath);

                // Check for dangerous extensions
                var dangerousExtensions = new[] { ".exe", ".bat", ".cmd", ".scr", ".pif", ".vbs", ".js", ".jar" };
                if (dangerousExtensions.Contains(extension))
                {
                    riskFactors.Add($"Potentially dangerous file extension: {extension}");
                    assessment.RiskLevel = SecurityRiskLevel.Critical;
                    assessment.IsBlocked = true;
                }

                // Check for double extensions
                var dotCount = fileName.Count(c => c == '.');
                                 if (dotCount > 1)
                {
                    riskFactors.Add("Multiple file extensions detected");
                    assessment.RiskLevel = (SecurityRiskLevel)Math.Max((int)assessment.RiskLevel, (int)SecurityRiskLevel.Medium);
                }

                // Check for suspicious patterns in filename
                var suspiciousPatterns = new[] { "..", "~", "javascript:", "vbscript:", "<script" };
                foreach (var pattern in suspiciousPatterns)
                {
                                         if (fileName.ToLowerInvariant().Contains(pattern))
                    {
                        riskFactors.Add($"Suspicious pattern in filename: {pattern}");
                        assessment.RiskLevel = (SecurityRiskLevel)Math.Max((int)assessment.RiskLevel, (int)SecurityRiskLevel.High);
                    }
                }

                                 // Check file size anomalies
                var maxSizeStr = _configService.Config.ModeDefaults.MaxFileSize ?? "100MB";
                var maxSize = ParseFileSizeToMB(maxSizeStr) * 1024L * 1024L;
                                 if (fileInfo.Length > maxSize)
                {
                    riskFactors.Add($"File size exceeds limit: {FormatFileSize(fileInfo.Length)}");
                    assessment.RiskLevel = (SecurityRiskLevel)Math.Max((int)assessment.RiskLevel, (int)SecurityRiskLevel.Medium);
                }

                // Check for zero-byte files
                                 if (fileInfo.Length == 0)
                {
                    riskFactors.Add("Zero-byte file detected");
                    assessment.RiskLevel = (SecurityRiskLevel)Math.Max((int)assessment.RiskLevel, (int)SecurityRiskLevel.Low);
                }

                // Path traversal check
                var normalizedPath = Path.GetFullPath(filePath);
                if (filePath != normalizedPath || filePath.Contains(".."))
                {
                    riskFactors.Add("Potential path traversal detected");
                    assessment.RiskLevel = SecurityRiskLevel.Critical;
                    assessment.IsBlocked = true;
                }

                assessment.RiskFactors = riskFactors;

                // Generate recommendations
                if (assessment.RiskLevel >= SecurityRiskLevel.High)
                {
                    assessment.Recommendation = "File blocked due to high security risk. Please verify file source and content.";
                }
                else if (assessment.RiskLevel == SecurityRiskLevel.Medium)
                {
                    assessment.Recommendation = "Proceed with caution. Verify file content before processing.";
                }
                else if (assessment.RiskLevel == SecurityRiskLevel.Low)
                {
                    assessment.Recommendation = "Minor security concerns detected. File can be processed safely.";
                }
                else
                {
                    assessment.Recommendation = "No security risks detected.";
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error assessing security risk for file: {FilePath}", filePath);
                assessment.RiskLevel = SecurityRiskLevel.Medium;
                assessment.RiskFactors.Add($"Error during security assessment: {ex.Message}");
                assessment.Recommendation = "Could not complete security assessment. Proceed with caution.";
            }

            return assessment;
        }

        public string SanitizeFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return string.Empty;

            try
            {
                // Remove dangerous characters
                var invalidChars = Path.GetInvalidPathChars().Concat(Path.GetInvalidFileNameChars()).ToArray();
                var sanitized = new string(filePath.Where(c => !invalidChars.Contains(c)).ToArray());

                // Remove relative path components
                sanitized = sanitized.Replace("..", "").Replace("~", "");

                // Normalize path separators
                sanitized = sanitized.Replace('/', Path.DirectorySeparatorChar);

                return sanitized;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error sanitizing file path: {FilePath}", filePath);
                return Path.GetFileName(filePath); // Fallback to just filename
            }
        }

        private async Task ValidateFileContentAsync(string filePath, EnhancedFileValidationResult result, CancellationToken cancellationToken)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                
                // Read first few bytes to check file signature
                using var stream = File.OpenRead(filePath);
                var buffer = new byte[8];
                var bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                
                if (bytesRead > 0)
                {
                    var isValidSignature = ValidateFileSignature(buffer, extension);
                    if (!isValidSignature)
                    {
                        result.Warnings.Add($"File signature does not match extension {extension}");
                    }
                }

                // Additional content validation based on file type
                switch (extension)
                {
                    case ".pdf":
                        await ValidatePdfContentAsync(filePath, result, cancellationToken);
                        break;
                    case ".docx":
                    case ".xlsx":
                        await ValidateOfficeXmlContentAsync(filePath, result, cancellationToken);
                        break;
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error validating file content: {FilePath}", filePath);
                result.Warnings.Add($"Content validation error: {ex.Message}");
            }
        }

        private bool ValidateFileSignature(byte[] buffer, string extension)
        {
            // PDF signature
            if (extension == ".pdf")
            {
                return buffer.Length >= 4 && 
                       buffer[0] == 0x25 && buffer[1] == 0x50 && buffer[2] == 0x44 && buffer[3] == 0x46; // %PDF
            }

            // ZIP-based Office files (docx, xlsx)
            if (extension == ".docx" || extension == ".xlsx")
            {
                return buffer.Length >= 2 && buffer[0] == 0x50 && buffer[1] == 0x4B; // PK
            }

            // Add more signature validations as needed
            return true; // Default to valid for unknown types
        }

        private async Task ValidatePdfContentAsync(string filePath, EnhancedFileValidationResult result, CancellationToken cancellationToken)
        {
            // Basic PDF validation - check if file is not corrupted
            try
            {
                using var stream = File.OpenRead(filePath);
                var buffer = new byte[1024];
                await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                
                var content = System.Text.Encoding.ASCII.GetString(buffer);
                if (!content.Contains("%PDF"))
                {
                    result.Warnings.Add("PDF file may be corrupted or invalid");
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"PDF validation failed: {ex.Message}");
            }
        }

        private async Task ValidateOfficeXmlContentAsync(string filePath, EnhancedFileValidationResult result, CancellationToken cancellationToken)
        {
            // Basic Office XML validation - check if it's a valid ZIP file
            try
            {
                using var stream = File.OpenRead(filePath);
                var buffer = new byte[22];
                await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken);
                
                // Check for ZIP signature
                if (buffer.Length < 2 || buffer[0] != 0x50 || buffer[1] != 0x4B)
                {
                    result.Warnings.Add("Office document may be corrupted or invalid");
                }
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"Office document validation failed: {ex.Message}");
            }
        }

        private Dictionary<string, string> InitializeMimeTypeMap()
        {
            return new Dictionary<string, string>
            {
                { ".pdf", "application/pdf" },
                { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
                { ".doc", "application/msword" },
                { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                { ".xls", "application/vnd.ms-excel" },
                { ".txt", "text/plain" },
                { ".rtf", "application/rtf" }
            };
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes < 1024) return $"{bytes} B";
            if (bytes < 1024 * 1024) return $"{bytes / 1024.0:F1} KB";
            if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024.0 * 1024.0):F1} MB";
            return $"{bytes / (1024.0 * 1024.0 * 1024.0):F1} GB";
        }

        private static int ParseFileSizeToMB(string sizeStr)
        {
            if (string.IsNullOrWhiteSpace(sizeStr))
                return 100; // Default 100MB

            try
            {
                var cleanStr = sizeStr.ToUpperInvariant().Trim();
                
                if (cleanStr.EndsWith("MB"))
                {
                    var value = cleanStr.Replace("MB", "").Trim();
                    return int.TryParse(value, out var mb) ? mb : 100;
                }
                else if (cleanStr.EndsWith("GB"))
                {
                    var value = cleanStr.Replace("GB", "").Trim();
                    return int.TryParse(value, out var gb) ? gb * 1024 : 100;
                }
                else if (cleanStr.EndsWith("KB"))
                {
                    var value = cleanStr.Replace("KB", "").Trim();
                    return int.TryParse(value, out var kb) ? Math.Max(1, kb / 1024) : 100;
                }
                else
                {
                    // Assume bytes
                    return int.TryParse(cleanStr, out var bytes) ? Math.Max(1, bytes / (1024 * 1024)) : 100;
                }
            }
            catch
            {
                return 100; // Default fallback
            }
        }

        // Legacy interface implementations for backward compatibility
        public async Task<LegacyValidationResult> ValidateAsync(IEnumerable<FileItem> files, CancellationToken cancellationToken = default)
        {
            var result = new LegacyValidationResult { IsValid = true };
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
                    cancellationToken.ThrowIfCancellationRequested();

                    var enhancedResult = await ValidateFileAsync(file.FilePath, cancellationToken);
                    
                    if (enhancedResult.IsValid)
                    {
                        result.ValidFiles.Add(file);
                        file.ValidationStatus = ValidationStatus.Valid;
                    }
                    else
                    {
                        result.InvalidFiles.Add(file);
                        file.ValidationStatus = ValidationStatus.Invalid;
                        file.ValidationError = string.Join("; ", enhancedResult.Errors);
                        
                        if (!string.IsNullOrEmpty(result.ErrorMessage))
                            result.ErrorMessage += "; ";
                        result.ErrorMessage += file.ValidationError;
                        
                        result.Warnings.AddRange(enhancedResult.Warnings);
                    }
                }

                result.IsValid = result.ValidFiles.Any() && result.InvalidFiles.Count == 0;
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Legacy validation failed");
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

            try
            {
                foreach (var filePath in filePaths)
                {
                    cancellationToken.ThrowIfCancellationRequested();

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

                    var enhancedResult = await ValidateFileAsync(filePath, cancellationToken);
                    
                    if (enhancedResult.IsValid)
                    {
                        fileItem.ValidationStatus = ValidationStatus.Valid;
                        fileItems.Add(fileItem);
                    }
                    else
                    {
                        fileItem.ValidationStatus = ValidationStatus.Invalid;
                        fileItem.ValidationError = string.Join("; ", enhancedResult.Errors);
                        _logger.Warning("Invalid file skipped: {FileName} - {Error}", 
                            fileItem.FileName, fileItem.ValidationError);
                    }
                }

                return fileItems;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Dropped file validation failed");
                return fileItems;
            }
        }

        public LegacyValidationResult ValidateQuick(IEnumerable<FileItem> files)
        {
            var result = new LegacyValidationResult { IsValid = true };
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
                    if (string.IsNullOrWhiteSpace(file.FilePath))
                    {
                        result.InvalidFiles.Add(file);
                        result.ErrorMessage = "File path is empty";
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(file.FileName))
                    {
                        result.InvalidFiles.Add(file);
                        result.ErrorMessage = "File name is empty";
                        continue;
                    }

                    if (IsFileTypeSupported(file.FilePath))
                    {
                        result.ValidFiles.Add(file);
                    }
                    else
                    {
                        result.InvalidFiles.Add(file);
                        result.ErrorMessage = $"File type not supported: {file.FileType}";
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
    }
} 