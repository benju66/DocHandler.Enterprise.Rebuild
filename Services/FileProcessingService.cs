using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Serilog;
using Task = System.Threading.Tasks.Task;

namespace DocHandler.Services
{
    public class FileProcessingService
    {
        private readonly ILogger _logger = Log.ForContext<FileProcessingService>();
        private readonly OfficeConversionService _officeConversionService;
        private readonly PdfOperationsService _pdfOperationsService;

        private readonly HashSet<string> _supportedExtensions = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".pdf", ".doc", ".docx", ".xls", ".xlsx"
        };

        public FileProcessingService()
        {
            _officeConversionService = new OfficeConversionService();
            _pdfOperationsService = new PdfOperationsService();
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

        public async Task<ProcessingResult> ProcessFiles(List<string> filePaths, string outputDirectory, bool convertOfficeToPdf = true)
        {
            var result = new ProcessingResult();
            _logger.Information("Processing {Count} files to {OutputDirectory}", filePaths.Count, outputDirectory);

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
                                foreach (var file in group)
                                {
                                    var outputPath = Path.Combine(outputDirectory, 
                                        Path.GetFileNameWithoutExtension(file) + ".pdf");
                                    
                                    var conversionResult = await _officeConversionService.ConvertWordToPdf(file, outputPath);
                                    if (conversionResult.Success)
                                    {
                                        processedPdfs.Add(outputPath);
                                        result.SuccessfulFiles.Add(outputPath);
                                        _logger.Information("Converted Word to PDF: {File}", Path.GetFileName(file));
                                    }
                                    else
                                    {
                                        result.FailedFiles.Add((file, conversionResult.ErrorMessage ?? "Unknown error"));
                                        result.ErrorMessage = conversionResult.ErrorMessage;
                                    }
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
                                foreach (var file in group)
                                {
                                    var outputPath = Path.Combine(outputDirectory, 
                                        Path.GetFileNameWithoutExtension(file) + ".pdf");
                                    
                                    var conversionResult = await _officeConversionService.ConvertExcelToPdf(file, outputPath);
                                    if (conversionResult.Success)
                                    {
                                        processedPdfs.Add(outputPath);
                                        result.SuccessfulFiles.Add(outputPath);
                                        _logger.Information("Converted Excel to PDF: {File}", Path.GetFileName(file));
                                    }
                                    else
                                    {
                                        result.FailedFiles.Add((file, conversionResult.ErrorMessage ?? "Unknown error"));
                                        result.ErrorMessage = conversionResult.ErrorMessage;
                                    }
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

                _logger.Information("Processing complete. Processed: {Processed}, Failed: {Failed}", 
                    result.SuccessfulFiles.Count, result.FailedFiles.Count);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during file processing");
                result.Success = false;
                result.ErrorMessage = $"Processing error: {ex.Message}";
                return result;
            }
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
    }

    public class ProcessingResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
        public List<string> SuccessfulFiles { get; set; } = new();
        public List<(string FilePath, string Error)> FailedFiles { get; set; } = new();
        public string OutputDirectory { get; set; }
        public bool IsMerged { get; set; }
    }
}