using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class FileProcessingService
    {
        private readonly ILogger _logger;
        private readonly List<string> _supportedExtensions = new() { ".pdf", ".doc", ".docx", ".xls", ".xlsx" };
        
        public FileProcessingService()
        {
            _logger = Log.ForContext<FileProcessingService>();
        }

        public bool IsFileSupported(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            return _supportedExtensions.Contains(extension);
        }

        public List<string> ValidateDroppedFiles(string[] files)
        {
            var validFiles = new List<string>();
            
            foreach (var file in files)
            {
                if (File.Exists(file) && IsFileSupported(file))
                {
                    validFiles.Add(file);
                    _logger.Information("Valid file added: {FilePath}", file);
                }
                else
                {
                    _logger.Warning("Invalid or unsupported file: {FilePath}", file);
                }
            }
            
            return validFiles;
        }

        public async Task<ProcessingResult> ProcessDroppedFiles(List<string> files, string outputPath)
        {
            var result = new ProcessingResult();
            
            try
            {
                // Create output directory if it doesn't exist
                Directory.CreateDirectory(outputPath);
                
                foreach (var file in files)
                {
                    try
                    {
                        await ProcessSingleFile(file, outputPath);
                        result.SuccessfulFiles.Add(file);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to process file: {FilePath}", file);
                        result.FailedFiles.Add((file, ex.Message));
                    }
                }
                
                result.Success = result.FailedFiles.Count == 0;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to process files");
                result.Success = false;
                result.ErrorMessage = ex.Message;
            }
            
            return result;
        }

        private async Task ProcessSingleFile(string filePath, string outputPath)
        {
            var fileName = Path.GetFileName(filePath);
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            // For now, just copy files. Later we'll add conversion logic
            var outputFileName = GetUniqueFileName(outputPath, fileName);
            var outputFilePath = Path.Combine(outputPath, outputFileName);
            
            // Use async file copy
            await Task.Run(() => File.Copy(filePath, outputFilePath, overwrite: false));
            
            _logger.Information("File processed: {InputFile} -> {OutputFile}", filePath, outputFilePath);
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
    }

    public class ProcessingResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> SuccessfulFiles { get; set; } = new();
        public List<(string FilePath, string Error)> FailedFiles { get; set; } = new();
    }
}