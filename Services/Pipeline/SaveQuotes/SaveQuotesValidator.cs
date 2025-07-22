using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocHandler.Models;
using DocHandler.Services.Pipeline;
using Serilog;

namespace DocHandler.Services.Pipeline.SaveQuotes
{
    /// <summary>
    /// Validator for Save Quotes mode - validates files can be processed for quote extraction
    /// </summary>
    public class SaveQuotesValidator : IFileValidator
    {
        private readonly IConfigurationService _configService;
        private readonly ILogger _logger;

        public string StageName => "SaveQuotes File Validation";

        private static readonly string[] SupportedExtensions = { ".doc", ".docx", ".pdf" };

        public SaveQuotesValidator(IConfigurationService configService)
        {
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _logger = Log.ForContext<SaveQuotesValidator>();
        }

        public async Task<bool> CanProcessAsync(FileItem file, ProcessingContext context)
        {
            // SaveQuotes validator can process Word documents and PDFs
            var extension = Path.GetExtension(file.FilePath).ToLowerInvariant();
            return SupportedExtensions.Contains(extension);
        }

        public async Task<ValidationResult> ValidateAsync(FileItem file, ProcessingContext context)
        {
            var result = new ValidationResult();

            try
            {
                _logger.Debug("Validating file for SaveQuotes processing: {FilePath}", file.FilePath);

                // Check if file exists
                if (!File.Exists(file.FilePath))
                {
                    result.ErrorMessages.Add($"File does not exist: {file.FileName}");
                    result.IsValid = false;
                    return result;
                }

                // Check file extension
                var extension = Path.GetExtension(file.FilePath).ToLowerInvariant();
                if (!SupportedExtensions.Contains(extension))
                {
                    result.ErrorMessages.Add($"Unsupported file type: {extension}. Supported types: {string.Join(", ", SupportedExtensions)}");
                    result.IsValid = false;
                    return result;
                }

                // Check file size
                var fileInfo = new FileInfo(file.FilePath);
                var maxSizeMB = _configService.Config.DocFileSizeLimitMB;
                var maxSizeBytes = maxSizeMB * 1024 * 1024;

                if (fileInfo.Length > maxSizeBytes)
                {
                    result.ErrorMessages.Add($"File size ({fileInfo.Length / (1024 * 1024):F1} MB) exceeds limit of {maxSizeMB} MB");
                    result.IsValid = false;
                    return result;
                }

                // Check if file is accessible (not locked)
                try
                {
                    using (var stream = File.OpenRead(file.FilePath))
                    {
                        // Just opening for read to check accessibility
                    }
                }
                catch (IOException ex)
                {
                    result.ErrorMessages.Add($"File is locked or inaccessible: {ex.Message}");
                    result.IsValid = false;
                    return result;
                }

                // Check for password protection (basic check)
                if (extension.StartsWith(".doc"))
                {
                    if (await IsPasswordProtectedAsync(file.FilePath))
                    {
                        result.Warnings.Add($"File may be password protected: {file.FileName}");
                        // Don't fail validation, but warn user
                    }
                }

                // Validation passed
                result.IsValid = true;
                result.Messages.Add($"File validated successfully: {file.FileName}");
                
                _logger.Debug("File validation successful: {FilePath}", file.FilePath);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error validating file: {FilePath}", file.FilePath);
                result.ErrorMessages.Add($"Validation error: {ex.Message}");
                result.IsValid = false;
                return result;
            }
        }

        private async Task<bool> IsPasswordProtectedAsync(string filePath)
        {
            try
            {
                // Basic check for password protection by trying to read the file
                // This is a simple heuristic - more sophisticated checks could be added
                using (var stream = File.OpenRead(filePath))
                {
                    var buffer = new byte[512];
                    await stream.ReadAsync(buffer, 0, buffer.Length);
                    
                    // Check for common password protection indicators in Office files
                    var content = System.Text.Encoding.ASCII.GetString(buffer);
                    return content.Contains("EncryptedPackage") || content.Contains("Microsoft Office Encrypted");
                }
            }
            catch
            {
                // If we can't read the file, it might be password protected
                return true;
            }
        }
    }
} 