using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Serilog;

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

        public class ValidationResult
        {
            public bool IsValid { get; set; }
            public string ErrorMessage { get; set; } = "";
            public List<string> Warnings { get; set; } = new List<string>();
            public FileInfo? FileInfo { get; set; }
            public bool IsSecure { get; set; } = true;
        }

        /// <summary>
        /// Comprehensive file validation with security checks
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

                // Path traversal protection
                if (HasPathTraversalAttempt(filePath))
                {
                    result.IsValid = false;
                    result.ErrorMessage = "File path contains invalid characters or path traversal attempts";
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

                // File signature validation (basic)
                if (!ValidateFileSignature(filePath, extension))
                {
                    result.Warnings.Add("File signature doesn't match expected format - file may be corrupted or renamed");
                }

                // Check for suspicious file characteristics
                CheckSuspiciousCharacteristics(filePath, result);

                result.IsValid = true;
                _logger.Debug("File validation passed: {File}", Path.GetFileName(filePath));
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
        /// Validates multiple files and returns results
        /// </summary>
        public static List<(string FilePath, ValidationResult Result)> ValidateFiles(IEnumerable<string> filePaths)
        {
            var results = new List<(string, ValidationResult)>();
            
            foreach (var filePath in filePaths)
            {
                var result = ValidateFile(filePath);
                results.Add((filePath, result));
            }

            return results;
        }

        /// <summary>
        /// Quick validation for dropped files
        /// </summary>
        public static bool IsValidDroppedFile(string filePath)
        {
            var result = ValidateFile(filePath);
            return result.IsValid;
        }

        /// <summary>
        /// Sanitizes filename to prevent security issues
        /// </summary>
        public static string SanitizeFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "Document";

            // Remove invalid characters
            var invalidChars = Path.GetInvalidFileNameChars();
            var sanitized = new StringBuilder();
            
            foreach (var c in fileName)
            {
                if (!invalidChars.Contains(c) && c != '<' && c != '>' && c != ':' && c != '"' && c != '|' && c != '?' && c != '*')
                {
                    sanitized.Append(c);
                }
                else
                {
                    sanitized.Append('_');
                }
            }

            var result = sanitized.ToString().Trim(' ', '.');
            
            // Ensure we don't have empty filename
            if (string.IsNullOrWhiteSpace(result))
                return "Document";

            // Limit length
            if (result.Length > 100)
                result = result.Substring(0, 100);

            return result;
        }

        private static bool HasPathTraversalAttempt(string filePath)
        {
            // Check for common path traversal patterns
            var suspiciousPatterns = new[] { "..", "~", "../", "..\\", "%2e%2e", "%2f", "%5c" };
            
            var normalizedPath = filePath.ToLowerInvariant();
            return suspiciousPatterns.Any(pattern => normalizedPath.Contains(pattern));
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
            }

            // Check for multiple extensions
            var dotCount = fileName.Count(c => c == '.');
            if (dotCount > 1)
            {
                result.Warnings.Add("Multiple file extensions detected");
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
    }
} 