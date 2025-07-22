using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocHandler.Models;
using DocHandler.Services.Pipeline;
using Serilog;

namespace DocHandler.Services.Pipeline.SaveQuotes
{
    /// <summary>
    /// Output generator for Save Quotes mode - finalizes organization and creates summary reports
    /// </summary>
    public class SaveQuotesOutputGenerator : IOutputGenerator
    {
        private readonly IConfigurationService _configService;
        private readonly ILogger _logger;

        public string StageName => "SaveQuotes Output Generation";

        public SaveQuotesOutputGenerator(IConfigurationService configService)
        {
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _logger = Log.ForContext<SaveQuotesOutputGenerator>();
        }

        public async Task<bool> CanProcessAsync(FileItem file, ProcessingContext context)
        {
            // Can generate output for SaveQuotes mode
            return context.Mode == ProcessingMode.SaveQuotes;
        }

        public async Task<OutputResult> GenerateAsync(List<PostProcessingResult> inputs, ProcessingContext context)
        {
            var result = new OutputResult
            {
                Success = true
            };

            try
            {
                _logger.Information("Generating final output for {FileCount} processed files", inputs.Count);

                // Step 1: Collect all successful outputs
                var successfulResults = inputs.Where(r => r.Success).ToList();
                var failedResults = inputs.Where(r => !r.Success).ToList();

                // Step 2: Organize final output paths
                foreach (var successfulResult in successfulResults)
                {
                    result.OutputPaths.Add(successfulResult.FinalPath);
                }

                // Step 3: Generate processing summary
                await GenerateProcessingSummaryAsync(successfulResults, failedResults, context, result);

                // Step 4: Update recent locations and configuration
                await UpdateRecentLocationsAsync(successfulResults, result);

                // Step 5: Create completion report
                await CreateCompletionReportAsync(successfulResults, failedResults, context, result);

                // Step 6: Add final metadata
                AddFinalMetadata(successfulResults, failedResults, context, result);

                result.Success = successfulResults.Any();
                result.Messages.Add($"Output generation completed: {successfulResults.Count} successful, {failedResults.Count} failed");

                _logger.Information("Output generation completed: {SuccessCount} successful, {FailedCount} failed", 
                    successfulResults.Count, failedResults.Count);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Output generation failed");
                result.Success = false;
                result.Error = ex;
                result.Messages.Add($"Output generation failed: {ex.Message}");
                return result;
            }
        }

        private async Task GenerateProcessingSummaryAsync(
            List<PostProcessingResult> successful, 
            List<PostProcessingResult> failed, 
            ProcessingContext context, 
            OutputResult result)
        {
            try
            {
                var summary = new ProcessingSummary
                {
                    ProcessingDate = DateTime.Now,
                    TotalFiles = successful.Count + failed.Count,
                    SuccessfulFiles = successful.Count,
                    FailedFiles = failed.Count,
                    OutputDirectory = context.OutputDirectory,
                    ProcessingMode = context.Mode.ToString()
                };

                // Analyze successful files
                foreach (var success in successful)
                {
                    var fileInfo = new ProcessedFileInfo
                    {
                        OriginalPath = success.SourceConversion.SourceFile.FilePath,
                        FinalPath = success.FinalPath,
                        CompanyName = success.SourceConversion.SourceFile.CompanyName,
                        ScopeOfWork = success.SourceConversion.SourceFile.ScopeOfWork,
                        ProcessingTime = success.SourceConversion.ProcessingTime,
                        FileSize = File.Exists(success.FinalPath) ? new FileInfo(success.FinalPath).Length : 0
                    };

                    summary.ProcessedFiles.Add(fileInfo);
                }

                // Analyze failed files
                foreach (var failure in failed)
                {
                    var errorInfo = new ProcessingErrorInfo
                    {
                        FilePath = failure.SourceConversion?.SourceFile?.FilePath ?? "Unknown",
                        ErrorMessage = failure.Error?.Message ?? "Unknown error",
                        Stage = "Post-Processing"
                    };

                    summary.Errors.Add(errorInfo);
                }

                result.OutputMetadata["ProcessingSummary"] = summary;
                result.Messages.Add($"Processing summary generated: {summary.SuccessfulFiles}/{summary.TotalFiles} files processed");

                _logger.Debug("Processing summary generated: {SuccessCount}/{TotalCount} files", summary.SuccessfulFiles, summary.TotalFiles);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to generate processing summary");
                result.Messages.Add($"Processing summary generation failed: {ex.Message}");
                // Don't fail the entire output generation for summary failure
            }
        }

        private async Task UpdateRecentLocationsAsync(List<PostProcessingResult> successful, OutputResult result)
        {
            try
            {
                // Extract unique directories from successful outputs
                var uniqueDirectories = successful
                    .Select(r => Path.GetDirectoryName(r.FinalPath))
                    .Where(dir => !string.IsNullOrEmpty(dir))
                    .Distinct()
                    .ToList();

                // TODO: Update configuration with recent locations when methods are available
                result.OutputMetadata["RecentLocationsUpdated"] = uniqueDirectories.Count;
                result.Messages.Add($"Found {uniqueDirectories.Count} unique output directories (recent locations update skipped)");
                
                _logger.Debug("Found {Count} unique directories", uniqueDirectories.Count);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to update recent locations");
                result.Messages.Add($"Recent locations update failed: {ex.Message}");
                // Don't fail the entire output generation for recent locations failure
            }
        }

        private async Task CreateCompletionReportAsync(
            List<PostProcessingResult> successful, 
            List<PostProcessingResult> failed, 
            ProcessingContext context, 
            OutputResult result)
        {
            try
            {
                // TODO: Check configuration for report generation when property is available
                // For now, always generate reports

                var reportFileName = $"SaveQuotes_Report_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                var reportPath = Path.Combine(context.OutputDirectory, reportFileName);

                var reportContent = BuildCompletionReport(successful, failed, context);

                await File.WriteAllTextAsync(reportPath, reportContent);

                result.OutputPaths.Add(reportPath);
                result.OutputMetadata["CompletionReportPath"] = reportPath;
                result.Messages.Add($"Completion report created: {reportFileName}");

                _logger.Information("Completion report created: {ReportPath}", reportPath);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to create completion report");
                result.Messages.Add($"Completion report creation failed: {ex.Message}");
                // Don't fail the entire output generation for report failure
            }
        }

        private string BuildCompletionReport(
            List<PostProcessingResult> successful, 
            List<PostProcessingResult> failed, 
            ProcessingContext context)
        {
            var report = new System.Text.StringBuilder();

            report.AppendLine("DocHandler Enterprise - SaveQuotes Processing Report");
            report.AppendLine("=".PadRight(60, '='));
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine($"Correlation ID: {context.CorrelationId}");
            report.AppendLine($"Output Directory: {context.OutputDirectory}");
            report.AppendLine();

            report.AppendLine("Summary:");
            report.AppendLine($"  Total Files: {successful.Count + failed.Count}");
            report.AppendLine($"  Successful: {successful.Count}");
            report.AppendLine($"  Failed: {failed.Count}");
            report.AppendLine($"  Success Rate: {(successful.Count / (double)(successful.Count + failed.Count) * 100):F1}%");
            report.AppendLine();

            if (successful.Any())
            {
                report.AppendLine("Successful Files:");
                report.AppendLine("-".PadRight(40, '-'));
                
                foreach (var success in successful)
                {
                    var sourceFile = success.SourceConversion.SourceFile;
                    report.AppendLine($"  File: {Path.GetFileName(sourceFile.FilePath)}");
                    report.AppendLine($"    Company: {sourceFile.CompanyName ?? "Not detected"}");
                    report.AppendLine($"    Scope: {sourceFile.ScopeOfWork ?? "Not assigned"}");
                    report.AppendLine($"    Output: {success.FinalPath}");
                    report.AppendLine($"    Processing Time: {success.SourceConversion.ProcessingTime.TotalSeconds:F1}s");
                    report.AppendLine();
                }
            }

            if (failed.Any())
            {
                report.AppendLine("Failed Files:");
                report.AppendLine("-".PadRight(40, '-'));
                
                foreach (var failure in failed)
                {
                    var sourceFile = failure.SourceConversion?.SourceFile;
                    report.AppendLine($"  File: {Path.GetFileName(sourceFile?.FilePath ?? "Unknown")}");
                    report.AppendLine($"    Error: {failure.Error?.Message ?? "Unknown error"}");
                    report.AppendLine();
                }
            }

            return report.ToString();
        }

        private void AddFinalMetadata(
            List<PostProcessingResult> successful, 
            List<PostProcessingResult> failed, 
            ProcessingContext context, 
            OutputResult result)
        {
            try
            {
                // Add comprehensive metadata about the processing session
                result.OutputMetadata["ProcessingCorrelationId"] = context.CorrelationId;
                result.OutputMetadata["ProcessingStartTime"] = DateTime.Now; // Should be from context
                result.OutputMetadata["ProcessingEndTime"] = DateTime.Now;
                result.OutputMetadata["TotalInputFiles"] = context.InputFiles.Count;
                result.OutputMetadata["SuccessfulOutputFiles"] = successful.Count;
                result.OutputMetadata["FailedFiles"] = failed.Count;
                result.OutputMetadata["OutputDirectory"] = context.OutputDirectory;
                result.OutputMetadata["ProcessingMode"] = context.Mode.ToString();

                // Calculate statistics
                if (successful.Any())
                {
                    var totalProcessingTime = successful.Sum(s => s.SourceConversion.ProcessingTime.TotalSeconds);
                    var averageProcessingTime = totalProcessingTime / successful.Count;
                    var totalOutputSize = successful
                        .Where(s => File.Exists(s.FinalPath))
                        .Sum(s => new FileInfo(s.FinalPath).Length);

                    result.OutputMetadata["TotalProcessingTimeSeconds"] = totalProcessingTime;
                    result.OutputMetadata["AverageProcessingTimeSeconds"] = averageProcessingTime;
                    result.OutputMetadata["TotalOutputSizeBytes"] = totalOutputSize;
                    result.OutputMetadata["TotalOutputSizeMB"] = totalOutputSize / (1024.0 * 1024.0);
                }

                // Company and scope statistics
                var companies = successful
                    .Select(s => s.SourceConversion.SourceFile.CompanyName)
                    .Where(c => !string.IsNullOrWhiteSpace(c))
                    .Distinct()
                    .ToList();

                var scopes = successful
                    .Select(s => s.SourceConversion.SourceFile.ScopeOfWork)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Distinct()
                    .ToList();

                result.OutputMetadata["UniqueCompanies"] = companies;
                result.OutputMetadata["UniqueScopes"] = scopes;
                result.OutputMetadata["CompanyCount"] = companies.Count;
                result.OutputMetadata["ScopeCount"] = scopes.Count;

                _logger.Debug("Final metadata added to output result");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to add final metadata");
                // Don't fail for metadata addition
            }
        }
    }

    /// <summary>
    /// Processing summary data structure
    /// </summary>
    public class ProcessingSummary
    {
        public DateTime ProcessingDate { get; set; }
        public int TotalFiles { get; set; }
        public int SuccessfulFiles { get; set; }
        public int FailedFiles { get; set; }
        public string OutputDirectory { get; set; }
        public string ProcessingMode { get; set; }
        public List<ProcessedFileInfo> ProcessedFiles { get; set; } = new();
        public List<ProcessingErrorInfo> Errors { get; set; } = new();
    }

    /// <summary>
    /// Information about a successfully processed file
    /// </summary>
    public class ProcessedFileInfo
    {
        public string OriginalPath { get; set; }
        public string FinalPath { get; set; }
        public string CompanyName { get; set; }
        public string ScopeOfWork { get; set; }
        public TimeSpan ProcessingTime { get; set; }
        public long FileSize { get; set; }
    }

    /// <summary>
    /// Information about a processing error
    /// </summary>
    public class ProcessingErrorInfo
    {
        public string FilePath { get; set; }
        public string ErrorMessage { get; set; }
        public string Stage { get; set; }
    }
} 