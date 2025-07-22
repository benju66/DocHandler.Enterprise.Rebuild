using System;
using System.IO;
using System.Threading.Tasks;
using DocHandler.Models;
using DocHandler.Services.Pipeline;
using Serilog;

namespace DocHandler.Services.Pipeline.SaveQuotes
{
    /// <summary>
    /// Converter for Save Quotes mode - processes files and generates quote documents
    /// </summary>
    public class SaveQuotesConverter : IFileConverter
    {
        private readonly IOptimizedFileProcessingService _fileProcessingService;
        private readonly ILogger _logger;

        public string StageName => "SaveQuotes Conversion";

        public SaveQuotesConverter(IOptimizedFileProcessingService fileProcessingService)
        {
            _fileProcessingService = fileProcessingService ?? throw new ArgumentNullException(nameof(fileProcessingService));
            _logger = Log.ForContext<SaveQuotesConverter>();
        }

        public async Task<bool> CanProcessAsync(FileItem file, ProcessingContext context)
        {
            // Can process files in SaveQuotes mode
            return context.Mode == ProcessingMode.SaveQuotes;
        }

        public async Task<ConversionResult> ConvertAsync(FileItem file, ProcessingContext context)
        {
            var result = new ConversionResult
            {
                SourceFile = file,
                Success = false
            };

            var startTime = DateTime.Now;

            try
            {
                _logger.Information("Converting file for SaveQuotes: {FilePath}", file.FilePath);

                // Get output directory and filename from pre-processing
                var outputDirectory = context.Properties.TryGetValue("OutputDirectory", out var outputDirObj)
                    ? outputDirObj.ToString()
                    : context.OutputDirectory;

                var organizedFileName = context.Properties.TryGetValue("OrganizedFileName", out var fileNameObj)
                    ? fileNameObj.ToString()
                    : Path.GetFileName(file.FilePath);

                var outputPath = Path.Combine(outputDirectory, organizedFileName);

                // Ensure output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Check if file is already a PDF
                var extension = Path.GetExtension(file.FilePath).ToLowerInvariant();
                if (extension == ".pdf")
                {
                    // For PDF files, just copy to organized location
                    await CopyFileToOutputAsync(file.FilePath, outputPath, result);
                }
                else if (extension == ".doc" || extension == ".docx")
                {
                    // Convert Word documents to PDF
                    await ConvertWordToPdfAsync(file, outputPath, result);
                }
                else
                {
                    throw new NotSupportedException($"File type {extension} is not supported for SaveQuotes conversion");
                }

                result.ProcessingTime = DateTime.Now - startTime;
                
                if (result.Success)
                {
                    _logger.Information("File conversion successful: {InputPath} -> {OutputPath}", 
                        file.FilePath, result.OutputPath);
                }

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Conversion failed for file: {FilePath}", file.FilePath);
                result.Success = false;
                result.Error = ex;
                result.Messages.Add($"Conversion failed: {ex.Message}");
                result.ProcessingTime = DateTime.Now - startTime;
                return result;
            }
        }

        private async Task CopyFileToOutputAsync(string inputPath, string outputPath, ConversionResult result)
        {
            try
            {
                // If source and destination are the same, no need to copy
                if (string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
                {
                    result.Success = true;
                    result.OutputPath = inputPath;
                    result.Messages.Add("File already in correct location");
                    return;
                }

                // Copy the file to the organized location
                File.Copy(inputPath, outputPath, overwrite: true);

                result.Success = true;
                result.OutputPath = outputPath;
                result.Messages.Add($"File copied to organized location: {Path.GetFileName(outputPath)}");
                result.ConversionData["OperationType"] = "Copy";

                _logger.Debug("File copied successfully: {InputPath} -> {OutputPath}", inputPath, outputPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to copy file: {InputPath} -> {OutputPath}", inputPath, outputPath);
                throw new InvalidOperationException($"Failed to copy file to output location: {ex.Message}", ex);
            }
        }

        private async Task ConvertWordToPdfAsync(FileItem file, string outputPath, ConversionResult result)
        {
            try
            {
                // Ensure output path has PDF extension
                var outputPdfPath = Path.ChangeExtension(outputPath, ".pdf");

                _logger.Debug("Converting Word document to PDF: {InputPath} -> {OutputPath}", file.FilePath, outputPdfPath);

                // Use the file processing service to convert to PDF
                var conversionResult = await _fileProcessingService.ProcessFileAsync(
                    file.FilePath,
                    outputPdfPath,
                    new ProgressCallback((current, total, fileName) =>
                    {
                        // Report progress if context has progress reporter
                        // This will be handled by the pipeline progress reporting
                    }));

                if (conversionResult)
                {
                    result.Success = true;
                    result.OutputPath = outputPdfPath;
                    result.Messages.Add($"Word document converted to PDF: {Path.GetFileName(outputPdfPath)}");
                    result.ConversionData["OperationType"] = "WordToPdf";
                    result.ConversionData["OriginalExtension"] = Path.GetExtension(file.FilePath);

                    _logger.Debug("Word to PDF conversion successful: {InputPath} -> {OutputPath}", file.FilePath, outputPdfPath);
                }
                else
                {
                    throw new InvalidOperationException("Word to PDF conversion failed - processing service returned false");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Word to PDF conversion failed: {InputPath}", file.FilePath);
                throw new InvalidOperationException($"Failed to convert Word document to PDF: {ex.Message}", ex);
            }
        }
    }
} 