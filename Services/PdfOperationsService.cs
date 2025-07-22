using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Serilog;
using Task = System.Threading.Tasks.Task;

namespace DocHandler.Services
{
    public class PdfOperationsService
    {
        private readonly ILogger _logger = Log.ForContext<PdfOperationsService>();

        public async Task<bool> MergePdfFiles(List<string> inputPaths, string outputPath)
        {
            _logger.Information("Starting PDF merge operation for {Count} files", inputPaths.Count);

            try
            {
                // THREADING FIX: Remove Task.Run wrapper - PDF operations should run on calling thread
                // Create output document
                using (var outputDocument = new PdfDocument())
                {
                    outputDocument.Info.Title = "Merged PDF Document";
                    outputDocument.Info.Author = "DocHandler";
                    outputDocument.Info.Creator = "DocHandler PDF Processor";

                    foreach (var inputPath in inputPaths)
                    {
                        try
                        {
                            _logger.Debug("Processing PDF: {Path}", inputPath);

                            // Open the document to copy pages from
                            using (var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import))
                            {
                                // Iterate through all pages
                                for (int idx = 0; idx < inputDocument.PageCount; idx++)
                                {
                                    // Get the page
                                    var page = inputDocument.Pages[idx];
                                    
                                    // Add the page to the output document
                                    outputDocument.AddPage(page);
                                }

                                _logger.Debug("Added {PageCount} pages from {FileName}", 
                                    inputDocument.PageCount, 
                                    Path.GetFileName(inputPath));
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Failed to process PDF: {Path}", inputPath);
                            throw new InvalidOperationException($"Failed to process {Path.GetFileName(inputPath)}: {ex.Message}", ex);
                        }
                    }

                    // Save the document
                    outputDocument.Save(outputPath);
                    _logger.Information("Successfully merged {Count} PDFs into {Output}", 
                        inputPaths.Count, 
                        Path.GetFileName(outputPath));

                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "PDF merge operation failed");
                return false;
            }
        }

        public async Task<int> GetPageCount(string pdfPath)
        {
            try
            {
                // THREADING FIX: Remove Task.Run wrapper - PDF operations should run on calling thread
                using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly))
                {
                    return document.PageCount;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to get page count for: {Path}", pdfPath);
                return 0;
            }
        }

        public async Task<bool> ValidatePdf(string pdfPath)
        {
            try
            {
                // THREADING FIX: Remove Task.Run wrapper - PDF operations should run on calling thread
                try
                {
                    using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly))
                    {
                        // If we can open it and it has at least one page, it's valid
                        return document.PageCount > 0;
                    }
                }
                catch
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "PDF validation failed for: {Path}", pdfPath);
                return false;
            }
        }

        public async Task<bool> CompressPdf(string inputPath, string outputPath, int compressionLevel = 5)
        {
            // TODO: Implement PDF compression
            // This will require additional libraries like iTextSharp or Ghostscript
            _logger.Warning("PDF compression not yet implemented");
            await Task.CompletedTask;
            return false;
        }

        public async Task<bool> SplitPdf(string inputPath, string outputDirectory, int pagesPerFile)
        {
            _logger.Information("Starting PDF split operation: {Path}", inputPath);

            try
            {
                // THREADING FIX: Remove Task.Run wrapper - PDF operations should run on calling thread
                using (var inputDocument = PdfReader.Open(inputPath, PdfDocumentOpenMode.Import))
                {
                    var totalPages = inputDocument.PageCount;
                    var fileCount = (int)Math.Ceiling((double)totalPages / pagesPerFile);
                    var baseFileName = Path.GetFileNameWithoutExtension(inputPath);

                    for (int fileIndex = 0; fileIndex < fileCount; fileIndex++)
                    {
                        using (var outputDocument = new PdfDocument())
                        {
                            var startPage = fileIndex * pagesPerFile;
                            var endPage = Math.Min(startPage + pagesPerFile, totalPages);

                            for (int pageIndex = startPage; pageIndex < endPage; pageIndex++)
                            {
                                outputDocument.AddPage(inputDocument.Pages[pageIndex]);
                            }

                            var outputFileName = $"{baseFileName}_part{fileIndex + 1}.pdf";
                            var outputPath = Path.Combine(outputDirectory, outputFileName);
                            outputDocument.Save(outputPath);

                            _logger.Debug("Created split file: {FileName} with {PageCount} pages", 
                                outputFileName, endPage - startPage);
                        }
                    }

                    _logger.Information("Successfully split PDF into {Count} files", fileCount);
                    return true;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "PDF split operation failed");
                return false;
            }
        }

        public async Task<PdfInfo> GetPdfInfo(string pdfPath)
        {
            try
            {
                // THREADING FIX: Remove Task.Run wrapper - PDF operations should run on calling thread
                using (var document = PdfReader.Open(pdfPath, PdfDocumentOpenMode.InformationOnly))
                {
                    var fileInfo = new FileInfo(pdfPath);
                    return new PdfInfo
                    {
                        FileName = Path.GetFileName(pdfPath),
                        FilePath = pdfPath,
                        PageCount = document.PageCount,
                        FileSize = fileInfo.Length,
                        Title = document.Info.Title ?? string.Empty,
                        Author = document.Info.Author ?? string.Empty,
                        Subject = document.Info.Subject ?? string.Empty,
                        Keywords = document.Info.Keywords ?? string.Empty,
                        Creator = document.Info.Creator ?? string.Empty,
                        Producer = document.Info.Producer ?? string.Empty,
                        CreationDate = document.Info.CreationDate,
                        ModificationDate = document.Info.ModificationDate
                    };
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to get PDF info for: {Path}", pdfPath);
                return null;
            }
        }
    }

    public class PdfInfo
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
        public int PageCount { get; set; }
        public long FileSize { get; set; }
        public string Title { get; set; }
        public string Author { get; set; }
        public string Subject { get; set; }
        public string Keywords { get; set; }
        public string Creator { get; set; }
        public string Producer { get; set; }
        public DateTime CreationDate { get; set; }
        public DateTime ModificationDate { get; set; }

        public string FileSizeFormatted => FormatFileSize(FileSize);

        private string FormatFileSize(long bytes)
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
    }
}