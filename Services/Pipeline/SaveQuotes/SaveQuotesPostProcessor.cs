using System;
using System.IO;
using System.Threading.Tasks;
using DocHandler.Models;
using DocHandler.Services.Pipeline;
using Serilog;
using System.Collections.Generic; // Added missing import for List

namespace DocHandler.Services.Pipeline.SaveQuotes
{
    /// <summary>
    /// Post-processor for Save Quotes mode - optimizes PDFs and adds metadata
    /// </summary>
    public class SaveQuotesPostProcessor : IPostProcessor
    {
        private readonly PdfOperationsService _pdfOperationsService;
        private readonly ILogger _logger;

        public string StageName => "SaveQuotes Post-Processing";

        public SaveQuotesPostProcessor(PdfOperationsService pdfOperationsService)
        {
            _pdfOperationsService = pdfOperationsService ?? throw new ArgumentNullException(nameof(pdfOperationsService));
            _logger = Log.ForContext<SaveQuotesPostProcessor>();
        }

        public async Task<bool> CanProcessAsync(FileItem file, ProcessingContext context)
        {
            // Can process PDF files in SaveQuotes mode
            return context.Mode == ProcessingMode.SaveQuotes;
        }

        public async Task<PostProcessingResult> ProcessAsync(ConversionResult input, ProcessingContext context)
        {
            var result = new PostProcessingResult
            {
                SourceConversion = input,
                Success = false,
                FinalPath = input.OutputPath
            };

            try
            {
                if (!input.Success)
                {
                    result.Success = false;
                    result.Messages.Add("Input conversion was not successful - skipping post-processing");
                    return result;
                }

                _logger.Information("Post-processing file: {FilePath}", input.OutputPath);

                // Only post-process PDF files
                var extension = Path.GetExtension(input.OutputPath).ToLowerInvariant();
                if (extension != ".pdf")
                {
                    result.Success = true;
                    result.Messages.Add("Non-PDF file - no post-processing required");
                    return result;
                }

                // Step 1: Basic file validation
                await ValidateFinalOutputAsync(result);

                // TODO: Add PDF optimization and metadata when PdfOperationsService methods are available
                result.Messages.Add("PDF optimization and metadata addition will be implemented when service methods are available");

                result.Success = true;
                result.Messages.Add($"Post-processing completed for {Path.GetFileName(input.OutputPath)}");
                
                _logger.Information("Post-processing successful for file: {FilePath}", input.OutputPath);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Post-processing failed for file: {FilePath}", input.OutputPath);
                result.Success = false;
                result.Error = ex;
                result.Messages.Add($"Post-processing failed: {ex.Message}");
                return result;
            }
        }



        private async Task ValidateFinalOutputAsync(PostProcessingResult result)
        {
            try
            {
                // Validate the final output file
                if (!File.Exists(result.FinalPath))
                {
                    throw new FileNotFoundException($"Final output file not found: {result.FinalPath}");
                }

                var fileInfo = new FileInfo(result.FinalPath);
                if (fileInfo.Length == 0)
                {
                    throw new InvalidOperationException($"Final output file is empty: {result.FinalPath}");
                }

                // Basic PDF validation if it's a PDF file (simplified validation)
                var extension = Path.GetExtension(result.FinalPath).ToLowerInvariant();
                if (extension == ".pdf")
                {
                    // Simple validation - check if file can be opened for reading
                    try
                    {
                        using (var stream = File.OpenRead(result.FinalPath))
                        {
                            var buffer = new byte[4];
                            var bytesRead = await stream.ReadAsync(buffer, 0, 4);
                            if (bytesRead < 4 || buffer[0] != 0x25 || buffer[1] != 0x50 || buffer[2] != 0x44 || buffer[3] != 0x46)
                            {
                                throw new InvalidOperationException($"File does not appear to be a valid PDF: {result.FinalPath}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Failed to validate PDF file: {ex.Message}", ex);
                    }
                }

                result.PostProcessingData["FileValidated"] = true;
                result.PostProcessingData["FinalFileSizeBytes"] = fileInfo.Length;
                result.Messages.Add($"Final output validated: {Path.GetFileName(result.FinalPath)}");
                
                _logger.Debug("Final output validation successful: {FilePath}", result.FinalPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Final output validation failed: {FilePath}", result.FinalPath);
                throw new InvalidOperationException($"Final output validation failed: {ex.Message}", ex);
            }
        }

    }
} 