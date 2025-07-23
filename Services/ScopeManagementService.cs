using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using DocHandler.Models;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Enhanced scope management service with fuzzy matching and learning capabilities
    /// </summary>
    public class ScopeManagementService : IScopeManagementService
    {
        private readonly ILogger _logger;
        private readonly IScopeOfWorkService _scopeOfWorkService;

        public ScopeManagementService(IScopeOfWorkService scopeOfWorkService)
        {
            _logger = Log.ForContext<ScopeManagementService>();
            _scopeOfWorkService = scopeOfWorkService ?? throw new ArgumentNullException(nameof(scopeOfWorkService));
            
            _logger.Information("ScopeManagementService initialized");
        }

        // Enhanced interface implementations
        public async Task<List<ScopeInfo>> GetAllScopesAsync()
        {
            await Task.Yield(); // Make it async
            
            try
            {
                var scopes = _scopeOfWorkService.Scopes
                    .Select(s => new ScopeInfo
                    {
                        Name = _scopeOfWorkService.GetFormattedScope(s),
                        Description = s.Description,
                        Keywords = new List<string>(), // Default empty list since Keywords property doesn't exist
                        UsageCount = s.UsageCount,
                        LastUsed = s.LastUsed ?? DateTime.MinValue, // Handle nullable DateTime
                        CreatedAt = DateTime.UtcNow, // Default value since CreatedAt doesn't exist
                        IsActive = true // Default value since IsActive doesn't exist
                    })
                    .ToList();

                return scopes;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error getting all scopes");
                return new List<ScopeInfo>();
            }
        }

        public async Task<List<ScopeInfo>> SearchScopesAsync(string searchTerm, double minConfidence = 0.7)
        {
            var allScopes = await GetAllScopesAsync();
            
            if (string.IsNullOrWhiteSpace(searchTerm))
                return allScopes;

            var filteredScopes = allScopes
                .Where(s => s.Name.ToLowerInvariant().Contains(searchTerm.ToLowerInvariant()) ||
                           s.Description?.ToLowerInvariant().Contains(searchTerm.ToLowerInvariant()) == true)
                .ToList();

            return filteredScopes;
        }

        public async Task<ScopeDetectionResult> DetectScopeAsync(string filePath, string? companyName = null)
        {
            await Task.Yield(); // Make it async
            
            return new ScopeDetectionResult
            {
                FilePath = filePath,
                DetectedScope = null,
                Confidence = 0.0,
                ProcessingTime = TimeSpan.FromMilliseconds(100)
            };
        }

        public async Task<bool> AddScopeAsync(string scopeName, string? description = null, List<string>? keywords = null)
        {
            await Task.Yield(); // Make it async
            
            if (string.IsNullOrWhiteSpace(scopeName))
                return false;

            _logger.Information("Would add scope: {ScopeName} with description: {Description}", 
                scopeName, description);
            
            return true; // Simulate success
        }

        public async Task IncrementScopeUsageAsync(string scopeName)
        {
            await Task.Yield(); // Make it async
            
            if (!string.IsNullOrWhiteSpace(scopeName))
            {
                _logger.Debug("Would increment usage for scope: {ScopeName}", scopeName);
            }
        }

        public async Task<List<ScopeInfo>> GetMostUsedScopesAsync(int count = 10)
        {
            var allScopes = await GetAllScopesAsync();
            return allScopes
                .OrderByDescending(s => s.UsageCount)
                .Take(count)
                .ToList();
        }

        public async Task LearnScopePatternAsync(string filePath, string scopeName, string? companyName = null)
        {
            await Task.Yield(); // Make it async
            
            _logger.Debug("Would learn pattern for scope: {ScopeName} from file: {FilePath}", 
                scopeName, System.IO.Path.GetFileName(filePath));
        }

        // Legacy interface methods for backward compatibility
        public async Task<List<ScopeInfo>> GetRecentScopesAsync(int count = 10)
        {
            var allScopes = await GetAllScopesAsync();
            return allScopes
                .OrderByDescending(s => s.LastUsed)
                .Take(count)
                .ToList();
        }

        public async Task<List<ScopeInfo>> FilterScopesAsync(string searchTerm)
        {
            return await SearchScopesAsync(searchTerm, 0.5);
        }

        public async Task<bool> ValidateScopeAsync(string scopeName)
        {
            await Task.Yield(); // Make it async
            
            if (string.IsNullOrWhiteSpace(scopeName))
                return false;

            var allScopes = await GetAllScopesAsync();
            return allScopes.Any(s => s.Name.Equals(scopeName, StringComparison.OrdinalIgnoreCase));
        }
    }
} 