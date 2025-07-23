using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using DocHandler.Models;
using DocHandler.Services.Pipeline;
using Serilog;
using System.Linq;

namespace DocHandler.Services.Pipeline.SaveQuotes
{
    /// <summary>
    /// Pre-processor for Save Quotes mode - extracts company names and assigns scopes
    /// </summary>
    public class SaveQuotesPreProcessor : IPreProcessor
    {
        private readonly IEnhancedCompanyDetectionService _companyDetectionService;
        private readonly IScopeManagementService _scopeManagementService;
        private readonly IConfigurationService _configService;
        private readonly ILogger _logger;

        public string StageName => "SaveQuotes Pre-Processing";

        public SaveQuotesPreProcessor(
            IEnhancedCompanyDetectionService companyDetectionService,
            IScopeManagementService scopeManagementService,
            IConfigurationService configService)
        {
            _companyDetectionService = companyDetectionService ?? throw new ArgumentNullException(nameof(companyDetectionService));
            _scopeManagementService = scopeManagementService ?? throw new ArgumentNullException(nameof(scopeManagementService));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _logger = Log.ForContext<SaveQuotesPreProcessor>();
        }

        public async Task<bool> CanProcessAsync(FileItem file, ProcessingContext context)
        {
            // Can process all files that passed validation for SaveQuotes
            return context.Mode == ProcessingMode.SaveQuotes;
        }

        public async Task<PreProcessingResult> ProcessAsync(FileItem file, ProcessingContext context)
        {
            var result = new PreProcessingResult
            {
                ProcessedFile = file,
                Success = true
            };

            try
            {
                _logger.Information("Pre-processing file for SaveQuotes: {FilePath}", file.FilePath);

                // Step 1: Detect company name if enabled and not already set
                if (string.IsNullOrWhiteSpace(file.CompanyName) && _configService.Config.AutoScanCompanyNames)
                {
                    await DetectCompanyNameAsync(file, context, result);
                }

                // Step 2: Determine or assign scope of work
                await AssignScopeOfWorkAsync(file, context, result);

                // Step 3: Validate we have required information
                if (string.IsNullOrWhiteSpace(file.CompanyName))
                {
                    result.ExtractedData["MissingCompanyName"] = true;
                    result.Messages.Add($"Warning: No company name detected for {file.FileName}");
                }

                if (string.IsNullOrWhiteSpace(file.ScopeOfWork))
                {
                    result.ExtractedData["MissingScopeOfWork"] = true;
                    result.Messages.Add($"Warning: No scope of work assigned for {file.FileName}");
                }

                // Step 4: Prepare output directory structure
                var outputInfo = await PrepareOutputDirectoryAsync(file, context, result);
                result.ExtractedData["OutputDirectory"] = outputInfo.OutputPath;
                result.ExtractedData["OrganizedFileName"] = outputInfo.OrganizedFileName;

                result.Messages.Add($"Pre-processing completed for {file.FileName}");
                _logger.Information("Pre-processing successful for file: {FilePath}", file.FilePath);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Pre-processing failed for file: {FilePath}", file.FilePath);
                result.Success = false;
                result.Error = ex;
                result.Messages.Add($"Pre-processing failed: {ex.Message}");
                return result;
            }
        }

        private async Task DetectCompanyNameAsync(FileItem file, ProcessingContext context, PreProcessingResult result)
        {
            try
            {
                _logger.Debug("Detecting company name for file: {FilePath}", file.FilePath);

                // Use the company detection service to scan the file
                var detectionResult = await _companyDetectionService.DetectCompanyAsync(file.FilePath);

                if (detectionResult != null && !string.IsNullOrWhiteSpace(detectionResult.DetectedCompany))
                {
                    file.CompanyName = detectionResult.DetectedCompany;
                    result.ExtractedData["DetectedCompanyName"] = detectionResult.DetectedCompany;
                    result.Messages.Add($"Company name detected: {detectionResult.DetectedCompany}");
                    
                    _logger.Information("Company name detected for {FilePath}: {CompanyName}", file.FilePath, detectionResult.DetectedCompany);
                }
                else
                {
                    result.ExtractedData["CompanyDetectionAttempted"] = true;
                    result.Messages.Add("No company name could be detected");
                    
                    _logger.Debug("No company name detected for file: {FilePath}", file.FilePath);
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Company name detection failed for file: {FilePath}", file.FilePath);
                result.ExtractedData["CompanyDetectionError"] = ex.Message;
                result.Messages.Add($"Company detection failed: {ex.Message}");
                // Don't fail the entire pre-processing for company detection failure
            }
        }

        private async Task AssignScopeOfWorkAsync(FileItem file, ProcessingContext context, PreProcessingResult result)
        {
            try
            {
                // If scope is already assigned, validate it exists
                if (!string.IsNullOrWhiteSpace(file.ScopeOfWork))
                {
                    var isValidScope = await _scopeManagementService.ValidateScopeAsync(file.ScopeOfWork);
                    if (isValidScope)
                    {
                        result.ExtractedData["ScopeValidated"] = true;
                        result.Messages.Add($"Scope of work validated: {file.ScopeOfWork}");
                        return;
                    }
                    else
                    {
                        result.Messages.Add($"Invalid scope of work: {file.ScopeOfWork}");
                        file.ScopeOfWork = null; // Clear invalid scope
                    }
                }

                // Try to auto-assign scope based on file name or content
                var searchTerm = ExtractScopeSearchTerm(file);
                if (!string.IsNullOrWhiteSpace(searchTerm))
                {
                    var suggestedScopes = await _scopeManagementService.FilterScopesAsync(searchTerm);
                    if (suggestedScopes.Any())
                    {
                        var topScope = suggestedScopes.First();
                        file.ScopeOfWork = topScope.Name;
                        result.ExtractedData["AutoAssignedScope"] = topScope.Name;
                        result.Messages.Add($"Scope auto-assigned: {topScope.Name}");
                        
                        _logger.Debug("Scope auto-assigned for {FilePath}: {Scope}", file.FilePath, topScope);
                    }
                }

                // If still no scope, record that manual assignment is needed
                if (string.IsNullOrWhiteSpace(file.ScopeOfWork))
                {
                    result.ExtractedData["RequiresManualScopeAssignment"] = true;
                    result.Messages.Add("Manual scope assignment required");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Scope assignment failed for file: {FilePath}", file.FilePath);
                result.ExtractedData["ScopeAssignmentError"] = ex.Message;
                result.Messages.Add($"Scope assignment failed: {ex.Message}");
                // Don't fail the entire pre-processing for scope assignment failure
            }
        }

        private string ExtractScopeSearchTerm(FileItem file)
        {
            try
            {
                // Extract potential scope terms from filename
                var fileName = Path.GetFileNameWithoutExtension(file.FilePath);
                
                // Look for common scope patterns in filename
                var patterns = new[]
                {
                    @"\b(\d{2}-\d{4})\b", // Pattern like "03-1000"
                    @"\b(scope|work|sow)\s*(\w+)", // "scope 1000", "work order", etc.
                    @"\b(\w+\s*\w*)\s*quote\b", // Something before "quote"
                };

                foreach (var pattern in patterns)
                {
                    var match = System.Text.RegularExpressions.Regex.Match(fileName, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        return match.Groups[1].Value.Trim();
                    }
                }

                // If no pattern found, use parts of the filename
                var words = fileName.Split(new[] { ' ', '_', '-' }, StringSplitOptions.RemoveEmptyEntries);
                if (words.Length > 0)
                {
                    return words[0]; // Use first word as search term
                }

                return string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        private async Task<OutputDirectoryInfo> PrepareOutputDirectoryAsync(FileItem file, ProcessingContext context, PreProcessingResult result)
        {
            var outputInfo = new OutputDirectoryInfo();

            try
            {
                // Build organized directory structure
                var baseOutputDir = context.OutputDirectory;
                var companyDir = !string.IsNullOrWhiteSpace(file.CompanyName) 
                    ? SanitizeDirectoryName(file.CompanyName)
                    : "Unknown Company";
                
                var scopeDir = !string.IsNullOrWhiteSpace(file.ScopeOfWork)
                    ? SanitizeDirectoryName(file.ScopeOfWork)
                    : "Unassigned";

                outputInfo.OutputPath = Path.Combine(baseOutputDir, companyDir, scopeDir);
                
                // Create organized filename
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var originalName = Path.GetFileNameWithoutExtension(file.FilePath);
                var extension = Path.GetExtension(file.FilePath);
                
                outputInfo.OrganizedFileName = $"{originalName}_{timestamp}{extension}";

                // Ensure directory exists
                Directory.CreateDirectory(outputInfo.OutputPath);

                result.Messages.Add($"Output directory prepared: {outputInfo.OutputPath}");

                return outputInfo;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to prepare output directory for file: {FilePath}", file.FilePath);
                
                // Fallback to base output directory
                outputInfo.OutputPath = context.OutputDirectory;
                outputInfo.OrganizedFileName = Path.GetFileName(file.FilePath);
                
                result.Messages.Add($"Using fallback output directory: {outputInfo.OutputPath}");
                return outputInfo;
            }
        }

        private string SanitizeDirectoryName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Default";

            // Remove invalid path characters
            var invalidChars = Path.GetInvalidPathChars().Concat(Path.GetInvalidFileNameChars()).ToArray();
            var sanitized = string.Join("_", name.Split(invalidChars, StringSplitOptions.RemoveEmptyEntries));
            
            // Trim and limit length
            sanitized = sanitized.Trim().Substring(0, Math.Min(sanitized.Length, 50));
            
            return string.IsNullOrWhiteSpace(sanitized) ? "Default" : sanitized;
        }

        private class OutputDirectoryInfo
        {
            public string OutputPath { get; set; }
            public string OrganizedFileName { get; set; }
        }
    }
} 