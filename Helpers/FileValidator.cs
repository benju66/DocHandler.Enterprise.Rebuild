using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Serilog;
using DocHandler.Services;

namespace DocHandler.Helpers
{
    public class FileValidator
    {
        private static readonly ILogger _logger = Log.ForContext<FileValidator>();

        // Maximum file size (50MB) - security measure
        private const long MAX_FILE_SIZE = 50 * 1024 * 1024;
        
        // Allowed file extensions for security
        private static readonly HashSet<string> AllowedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".pdf", ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".txt", ".rtf"
        };

        // File signatures (magic numbers) for basic file type validation
        private static readonly Dictionary<string, byte[]> FileSignatures = new Dictionary<string, byte[]>
        {
            { ".pdf", new byte[] { 0x25, 0x50, 0x44, 0x46 } }, // %PDF
            { ".docx", new byte[] { 0x50, 0x4B, 0x03, 0x04 } }, // ZIP signature (Office files)
            { ".doc", new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 } }, // OLE signature
            { ".xlsx", new byte[] { 0x50, 0x4B, 0x03, 0x04 } }, // ZIP signature
            { ".xls", new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 } }, // OLE signature
            { ".pptx", new byte[] { 0x50, 0x4B, 0x03, 0x04 } }, // ZIP signature
            { ".ppt", new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 } }, // OLE signature
            { ".txt", new byte[] { } }, // No specific signature for text files
            { ".rtf", new byte[] { 0x7B, 0x5C, 0x72, 0x74, 0x66 } } // {\rtf
        };

        // Enhanced security patterns
        private static readonly string[] SuspiciousPatterns = new[]
        {
            "..", "~", "../", "..\\", "%2e%2e", "%2f", "%5c", 
            "..%2f", "..%5c", "%2e%2e%2f", "%2e%2e%5c",
            "javascript:", "vbscript:", "data:", "file:",
            "<script", "</script", "eval(", "document.write"
        };

        private static readonly string[] DangerousExtensions = new[]
        {
            ".exe", ".bat", ".cmd", ".com", ".scr", ".pif", 
            ".vbs", ".js", ".jar", ".app", ".deb", ".pkg", 
            ".dmg", ".iso", ".msi", ".ps1", ".psm1"
        };

        public class ValidationResult
        {
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; } = "";
            public List<string> Warnings { get; set; } = new List<string>();
            public FileInfo? FileInfo { get; set; }
            public bool IsSecure { get; set; } = true;
            public SecurityRiskLevel RiskLevel { get; set; } = SecurityRiskLevel.None;
            public List<string> SecurityConcerns { get; set; } = new List<string>();
        }

        public enum SecurityRiskLevel
        {
            None,
            Low,
            Medium,
            High,
            Critical
        }

        /// <summary>
        /// Comprehensive file validation with enhanced security checks
        /// </summary>
        public static ValidationResult ValidateFile(string filePath)
        {
            var result = new ValidationResult();
            
            try
            {
                // Basic existence check
                if (!File.Exists(filePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "File does not exist";
                    return result;
                }

                var fileInfo = new FileInfo(filePath);
                result.FileInfo = fileInfo;

                // Enhanced security checks first
                var securityResult = PerformSecurityValidation(filePath, fileInfo);
                result.IsSecure = securityResult.IsSecure;
                result.RiskLevel = securityResult.RiskLevel;
                result.SecurityConcerns.AddRange(securityResult.Concerns);

                // Block high/critical security risks immediately
                if (result.RiskLevel >= SecurityRiskLevel.High)
                {
                    result.IsValid = false;
                    result.ErrorMessage = "File blocked due to security concerns";
                    _logger.Warning("File blocked due to security concerns: {File}, Risk: {RiskLevel}, Concerns: {Concerns}", 
                        filePath, result.RiskLevel, string.Join(", ", result.SecurityConcerns));
                    return result;
                }

                // File size check
                if (fileInfo.Length > MAX_FILE_SIZE)
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File size ({FormatFileSize(fileInfo.Length)}) exceeds maximum allowed size ({FormatFileSize(MAX_FILE_SIZE)})";
                    result.IsSecure = false;
                    return result;
                }

                // Empty file check
                if (fileInfo.Length == 0)
                {
                    result.IsValid = false;
                    result.ErrorMessage = "File is empty";
                    return result;
                }

                // Extension validation
                var extension = fileInfo.Extension.ToLowerInvariant();
                if (!AllowedExtensions.Contains(extension))
                {
                    result.IsValid = false;
                    result.ErrorMessage = $"File type '{extension}' is not allowed";
                    result.IsSecure = false;
                    return result;
                }

                // File access check
                if (!CanAccessFile(filePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "Cannot access file - it may be locked or you don't have permission";
                    return result;
                }

                // File signature validation (enhanced)
                if (!ValidateFileSignature(filePath, extension))
                {
                    result.RiskLevel = SecurityRiskLevel.Medium;
                    result.SecurityConcerns.Add("File signature mismatch - possible file type spoofing");
                    result.Warnings.Add("File signature doesn't match expected format - file may be corrupted or renamed");
                }

                // Enhanced content scanning for Office files
                if (IsOfficeFile(extension))
                {
                    var contentResult = ScanOfficeFileContent(filePath);
                    if (contentResult.HasMacros)
                    {
                        result.RiskLevel = GetMaxRiskLevel(result.RiskLevel, SecurityRiskLevel.Medium);
                        result.SecurityConcerns.Add("File contains macros");
                        result.Warnings.Add("File contains macros which may pose security risks");
                    }
                }

                // Check for suspicious file characteristics
                CheckSuspiciousCharacteristics(filePath, result);

                result.IsValid = true;
                _logger.Debug("File validation passed: {File}, Risk Level: {RiskLevel}", 
                    Path.GetFileName(filePath), result.RiskLevel);
                
                // Log security concerns if any
                if (result.SecurityConcerns.Any())
                {
                    _logger.Warning("File validation concerns for {File}: {Concerns}", 
                        Path.GetFileName(filePath), string.Join(", ", result.SecurityConcerns));
                }
            }
            catch (Exception ex)
            {
                result.IsValid = false;
                result.ErrorMessage = $"Validation error: {ex.Message}";
                result.IsSecure = false;
                _logger.Warning(ex, "File validation failed: {File}", filePath);
            }

            return result;
        }

        /// <summary>
        /// Validates a file and throws appropriate custom exceptions on failure
        /// </summary>
        public static void ValidateFileOrThrow(string filePath)
        {
            var result = ValidateFile(filePath);
            
            if (!result.IsValid)
            {
                if (!result.IsSecure || result.RiskLevel >= SecurityRiskLevel.High)
                {
                    if (HasPathTraversalAttempt(filePath))
                    {
                        throw ExceptionFactory.PathTraversal(filePath);
                    }
                    throw new SecurityViolationException("File Security", filePath, 
                        string.Join(", ", result.SecurityConcerns));
                }
                
                if (!File.Exists(filePath))
                {
                    throw ExceptionFactory.FileNotFound(filePath);
                }
                
                if (result.FileInfo?.Length > MAX_FILE_SIZE)
                {
                    throw ExceptionFactory.FileTooLarge(filePath, result.FileInfo.Length, MAX_FILE_SIZE);
                }
                
                var extension = Path.GetExtension(filePath);
                if (!AllowedExtensions.Contains(extension))
                {
                    throw ExceptionFactory.UnsupportedFileType(filePath, extension);
                }
                
                // Generic file validation exception for other cases
                var reason = DetermineValidationFailureReason(result);
                throw new FileValidationException(filePath, reason, result.ErrorMessage);
            }
        }

        /// <summary>
        /// Enhanced security validation with detailed risk assessment
        /// </summary>
        private static (bool IsSecure, SecurityRiskLevel RiskLevel, List<string> Concerns) PerformSecurityValidation(string filePath, FileInfo fileInfo)
        {
            var concerns = new List<string>();
            var riskLevel = SecurityRiskLevel.None;
            
            // Path traversal check
            if (HasPathTraversalAttempt(filePath))
            {
                concerns.Add("Path traversal attempt detected");
                riskLevel = SecurityRiskLevel.Critical;
            }
            
            // Dangerous extension check
            var extension = fileInfo.Extension.ToLowerInvariant();
            if (DangerousExtensions.Contains(extension))
            {
                concerns.Add($"Dangerous file extension: {extension}");
                riskLevel = SecurityRiskLevel.Critical;
            }
            
            // Double extension check
            var fileName = fileInfo.Name;
            var dotCount = fileName.Count(c => c == '.');
            if (dotCount > 1 && HasDangerousDoubleExtension(fileName))
            {
                concerns.Add("Suspicious double extension detected");
                riskLevel = GetMaxRiskLevel(riskLevel, SecurityRiskLevel.High);
            }
            
            // Suspicious filename patterns
            var suspiciousPatterns = SuspiciousPatterns.Where(pattern => 
                fileName.ToLowerInvariant().Contains(pattern.ToLowerInvariant())).ToList();
            
            if (suspiciousPatterns.Any())
            {
                concerns.Add($"Suspicious filename patterns: {string.Join(", ", suspiciousPatterns)}");
                riskLevel = GetMaxRiskLevel(riskLevel, SecurityRiskLevel.Medium);
            }
            
            // File size anomalies
            if (fileInfo.Length > MAX_FILE_SIZE)
            {
                concerns.Add("File size exceeds security limits");
                riskLevel = GetMaxRiskLevel(riskLevel, SecurityRiskLevel.Medium);
            }
            
            // Very small files that claim to be complex formats
            if (fileInfo.Length < 100 && (extension == ".docx" || extension == ".xlsx" || extension == ".pptx"))
            {
                concerns.Add("File too small for claimed format");
                riskLevel = GetMaxRiskLevel(riskLevel, SecurityRiskLevel.Medium);
            }
            
            var isSecure = riskLevel < SecurityRiskLevel.High;
            return (isSecure, riskLevel, concerns);
        }

        /// <summary>
        /// Scans Office files for potentially dangerous content
        /// </summary>
        private static (bool HasMacros, bool HasExternalLinks, List<string> Concerns) ScanOfficeFileContent(string filePath)
        {
            var concerns = new List<string>();
            var hasMacros = false;
            var hasExternalLinks = false;
            
            try
            {
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                
                // For macro-enabled formats, assume macros are present
                if (extension.EndsWith("m")) // .docm, .xlsm, .pptm
                {
                    hasMacros = true;
                    concerns.Add("Macro-enabled file format");
                }
                
                // Basic content scanning for modern Office files (ZIP-based)
                if (extension == ".docx" || extension == ".xlsx" || extension == ".pptx")
                {
                    using (var stream = File.OpenRead(filePath))
                    {
                        var buffer = new byte[1024];
                        var bytesRead = stream.Read(buffer, 0, buffer.Length);
                        var content = Encoding.UTF8.GetString(buffer, 0, bytesRead).ToLowerInvariant();
                        
                        // Look for macro-related content
                        if (content.Contains("vba") || content.Contains("macro") || content.Contains("activex"))
                        {
                            hasMacros = true;
                            concerns.Add("Potential macro content detected");
                        }
                        
                        // Look for external links
                        if (content.Contains("http://") || content.Contains("https://") || content.Contains("ftp://"))
                        {
                            hasExternalLinks = true;
                            concerns.Add("External links detected");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Debug(ex, "Could not scan file content for {File}", filePath);
                concerns.Add("Could not perform content security scan");
            }
            
            return (hasMacros, hasExternalLinks, concerns);
        }

        private static bool IsOfficeFile(string extension)
        {
            var officeExtensions = new[] { ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt" };
            return officeExtensions.Contains(extension);
        }

        private static bool HasDangerousDoubleExtension(string fileName)
        {
            // Check for patterns like document.pdf.exe
            var parts = fileName.Split('.');
            if (parts.Length >= 3)
            {
                var secondToLast = $".{parts[^2].ToLowerInvariant()}";
                var last = $".{parts[^1].ToLowerInvariant()}";
                
                // Safe extension followed by dangerous extension
                return AllowedExtensions.Contains(secondToLast) && DangerousExtensions.Contains(last);
            }
            return false;
        }

        private static ValidationFailureReason DetermineValidationFailureReason(ValidationResult result)
        {
            if (result.FileInfo == null) return ValidationFailureReason.Unknown;
            
            if (!File.Exists(result.FileInfo.FullName))
                return ValidationFailureReason.FileNotFound;
            
            if (result.FileInfo.Length == 0)
                return ValidationFailureReason.EmptyFile;
            
            if (result.FileInfo.Length > MAX_FILE_SIZE)
                return ValidationFailureReason.FileTooLarge;
            
            if (!result.IsSecure)
                return ValidationFailureReason.SecurityViolation;
            
            return ValidationFailureReason.Unknown;
        }

        /// <summary>
        /// Validates multiple files and returns results with enhanced error information
        /// </summary>
        public static List<(string FilePath, ValidationResult Result)> ValidateFiles(IEnumerable<string> filePaths)
        {
            var results = new List<(string, ValidationResult)>();
            
            foreach (var filePath in filePaths)
            {
                var result = ValidateFile(filePath);
                results.Add((filePath, result));
                
                // Log security concerns for each file
                if (result.SecurityConcerns.Any())
                {
                    _logger.Warning("Security concerns for {File}: {Concerns}", 
                        Path.GetFileName(filePath), string.Join(", ", result.SecurityConcerns));
                }
            }

            return results;
        }

        /// <summary>
        /// Quick validation for dropped files with enhanced security
        /// </summary>
        public static bool IsValidDroppedFile(string filePath)
        {
            var result = ValidateFile(filePath);
            return result.IsValid && result.IsSecure && result.RiskLevel < SecurityRiskLevel.High;
        }

        /// <summary>
        /// Enhanced filename sanitization with security focus
        /// </summary>
        public static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "Document";

            // Remove invalid characters and dangerous patterns
            var invalidChars = Path.GetInvalidFileNameChars().Concat(new[] { '<', '>', ':', '"', '|', '?', '*' }).ToArray();
            var sanitized = new StringBuilder();
            
            foreach (var c in fileName)
            {
                if (!invalidChars.Contains(c) && !char.IsControl(c))
                {
                    sanitized.Append(c);
                }
                else
                {
                    sanitized.Append('_');
                }
            }

            var result = sanitized.ToString().Trim(' ', '.');
            
            // Remove dangerous patterns
            foreach (var pattern in SuspiciousPatterns)
            {
                result = result.Replace(pattern, "_", StringComparison.OrdinalIgnoreCase);
            }
            
            // Ensure we don't have empty filename
            if (string.IsNullOrWhiteSpace(result))
                return "Document";

            // Limit length
            if (result.Length > 100)
                result = result.Substring(0, 100);

            return result;
        }

        /// <summary>
        /// Enhanced file path validation with security focus and path traversal prevention
        /// </summary>
        public static bool IsValidFilePath(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            try
            {
                // Get the full normalized path to prevent path traversal
                var fullPath = Path.GetFullPath(filePath);
                
                // Check if file exists
                if (!File.Exists(fullPath))
                    return false;

                // Enhanced path traversal prevention
                var allowedDirectories = GetAllowedDirectories();
                if (!allowedDirectories.Any(dir => fullPath.StartsWith(dir, StringComparison.OrdinalIgnoreCase)))
                {
                    _logger.Warning("File path outside allowed directories: {Path}", fullPath);
                    return false;
                }

                // Additional security checks
                if (HasPathTraversalAttempt(filePath))
                {
                    _logger.Warning("Path traversal attempt detected: {Path}", filePath);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "File path validation failed for: {Path}", filePath);
                return false;
            }
        }

        /// <summary>
        /// Gets the list of allowed directories for file operations
        /// </summary>
        private static List<string> GetAllowedDirectories()
        {
            var allowedDirs = new List<string>();

            try
            {
                // Add common safe directories
                allowedDirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
                allowedDirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.Desktop));
                allowedDirs.Add(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory));
                allowedDirs.Add(Path.GetTempPath());
                
                // Add application-specific directories
                var appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DocHandler");
                if (Directory.Exists(appDataPath))
                    allowedDirs.Add(appDataPath);

                // Add current working directory and subdirectories
                var currentDir = Directory.GetCurrentDirectory();
                allowedDirs.Add(currentDir);

                // Normalize all paths to full paths
                return allowedDirs.Select(Path.GetFullPath).Distinct().ToList();
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to build allowed directories list");
                // Return at least the current directory as fallback
                return new List<string> { Directory.GetCurrentDirectory() };
            }
        }

        /// <summary>
        /// Enhanced file extension validation with whitelist approach
        /// </summary>
        public static bool HasValidExtension(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                return false;

            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return AllowedExtensions.Contains(extension);
        }

        #region Private Helper Methods

        /// <summary>
        /// Helper method to get the maximum risk level between two values
        /// </summary>
        private static SecurityRiskLevel GetMaxRiskLevel(SecurityRiskLevel current, SecurityRiskLevel newLevel)
        {
            return (SecurityRiskLevel)Math.Max((int)current, (int)newLevel);
        }

        private static bool HasPathTraversalAttempt(string filePath)
        {
            // Enhanced path traversal detection
            var normalizedPath = filePath.ToLowerInvariant();
            return SuspiciousPatterns.Any(pattern => normalizedPath.Contains(pattern));
        }

        private static bool CanAccessFile(string filePath)
        {
            try
            {
                using (var stream = File.OpenRead(filePath))
                {
                    // Just open and close to test accessibility
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool ValidateFileSignature(string filePath, string extension)
        {
            if (!FileSignatures.ContainsKey(extension))
                return true; // No signature to validate

            var expectedSignature = FileSignatures[extension];
            if (expectedSignature.Length == 0)
                return true; // No signature requirement

            try
            {
                using (var stream = File.OpenRead(filePath))
                {
                    var buffer = new byte[expectedSignature.Length];
                    var bytesRead = stream.Read(buffer, 0, buffer.Length);
                    
                    if (bytesRead < expectedSignature.Length)
                        return false;

                    return buffer.SequenceEqual(expectedSignature);
                }
            }
            catch
            {
                return false;
            }
        }

        private static void CheckSuspiciousCharacteristics(string filePath, ValidationResult result)
        {
            var fileName = Path.GetFileName(filePath);
            
            // Check for suspicious filename patterns
            if (fileName.Contains(".."))
            {
                result.Warnings.Add("Filename contains suspicious patterns");
                result.SecurityConcerns.Add("Suspicious filename patterns");
                result.RiskLevel = GetMaxRiskLevel(result.RiskLevel, SecurityRiskLevel.Medium);
            }

            // Check for hidden files
            if (fileName.StartsWith("."))
            {
                result.Warnings.Add("Hidden file detected");
            }

            // Check for very long filenames
            if (fileName.Length > 100)
            {
                result.Warnings.Add("Filename is unusually long");
                result.SecurityConcerns.Add("Unusually long filename");
                result.RiskLevel = GetMaxRiskLevel(result.RiskLevel, SecurityRiskLevel.Low);
            }

            // Check for multiple extensions
            var dotCount = fileName.Count(c => c == '.');
            if (dotCount > 1)
            {
                result.Warnings.Add("Multiple file extensions detected");
                if (HasDangerousDoubleExtension(fileName))
                {
                    result.SecurityConcerns.Add("Dangerous double extension");
                    result.RiskLevel = GetMaxRiskLevel(result.RiskLevel, SecurityRiskLevel.High);
                }
            }
        }

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            int order = 0;
            double size = bytes;
            
            while (size >= 1024 && order < sizes.Length - 1)
            {
                order++;
                size /= 1024;
            }
            
            return $"{size:0.##} {sizes[order]}";
        }

        #endregion
    }
} 