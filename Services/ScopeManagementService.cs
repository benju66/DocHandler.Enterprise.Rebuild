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
    /// Handles scope management logic extracted from MainViewModel
    /// </summary>
    public class ScopeManagementService : IScopeManagementService
    {
        private readonly ILogger _logger;
        private readonly ScopeOfWorkService _scopeService;
        private readonly IConfigurationService _configService;
        private readonly Dictionary<string, (string code, string description)> _scopePartsCache;

        public ScopeManagementService(
            ScopeOfWorkService scopeService,
            IConfigurationService configService)
        {
            _logger = Log.ForContext<ScopeManagementService>();
            _scopeService = scopeService ?? throw new ArgumentNullException(nameof(scopeService));
            _configService = configService ?? throw new ArgumentNullException(nameof(configService));
            _scopePartsCache = new Dictionary<string, (string code, string description)>();
            
            _logger.Debug("ScopeManagementService initialized");
        }

        public async Task<List<string>> SearchScopesAsync(ScopeSearchRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrWhiteSpace(request.SearchTerm))
            {
                _logger.Debug("Empty search term provided, returning all scopes");
                return await GetAllScopesAsync();
            }

            _logger.Debug("Searching scopes for term: {SearchTerm}", request.SearchTerm);

            try
            {
                if (request.FuzzySearch)
                {
                    return await PerformFuzzySearchAsync(request.SearchTerm, request.MaxResults);
                }
                else
                {
                    return await PerformExactSearchAsync(request.SearchTerm, request.MaxResults);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Scope search failed for term: {SearchTerm}", request.SearchTerm);
                return new List<string>();
            }
        }

        public async Task<List<string>> FilterScopesAsync(string searchTerm)
        {
            try
            {
                var request = new ScopeSearchRequest(searchTerm ?? string.Empty, 50, true);
                return await SearchScopesAsync(request);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Scope filtering failed for term: {SearchTerm}", searchTerm);
                return new List<string>();
            }
        }

        public async Task<List<string>> GetRecentScopesAsync()
        {
            try
            {
                await Task.CompletedTask; // Make async for consistency
                return _scopeService.RecentScopes.Take(10).ToList();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to get recent scopes");
                return new List<string>();
            }
        }

        public async Task UpdateScopeUsageAsync(string scope)
        {
            if (string.IsNullOrWhiteSpace(scope))
            {
                _logger.Warning("Cannot update scope usage: scope is empty");
                return;
            }

            try
            {
                _logger.Debug("Updating scope usage for: {Scope}", scope);
                await _scopeService.UpdateRecentScope(scope);
                await _scopeService.IncrementUsageCount(scope);
                _logger.Debug("Scope usage updated successfully for: {Scope}", scope);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to update scope usage for: {Scope}", scope);
            }
        }

        public async Task ClearRecentScopesAsync()
        {
            try
            {
                _logger.Information("Clearing recent scopes");
                await _scopeService.ClearRecentScopes();
                _logger.Information("Recent scopes cleared successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to clear recent scopes");
                throw;
            }
        }

        private async Task<List<string>> GetAllScopesAsync()
        {
            try
            {
                await Task.CompletedTask; // Make async for consistency
                return _scopeService.Scopes
                    .Select(s => _scopeService.GetFormattedScope(s))
                    .ToList();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to get all scopes");
                return new List<string>();
            }
        }

        private async Task<List<string>> PerformFuzzySearchAsync(string searchTerm, int maxResults)
        {
            try
            {
                _logger.Debug("Performing fuzzy search for: {SearchTerm}", searchTerm);

                // Get all scopes
                var allScopes = await GetAllScopesAsync();

                // Build filtered list using fuzzy search
                var searchWords = searchTerm.ToLowerInvariant()
                    .Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);

                var scoredScopes = new List<(string scope, double score)>();

                foreach (var scope in allScopes)
                {
                    var score = CalculateFuzzyScore(scope, searchTerm, searchWords);
                    if (score > 0)
                    {
                        scoredScopes.Add((scope, score));
                    }
                }

                // Sort by score and return top results
                var results = scoredScopes
                    .OrderByDescending(x => x.score)
                    .Take(maxResults)
                    .Select(x => x.scope)
                    .ToList();

                _logger.Debug("Fuzzy search completed: {ResultCount} results for {SearchTerm}", 
                    results.Count, searchTerm);

                return results;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Fuzzy search failed for: {SearchTerm}", searchTerm);
                return new List<string>();
            }
        }

        private async Task<List<string>> PerformExactSearchAsync(string searchTerm, int maxResults)
        {
            try
            {
                _logger.Debug("Performing exact search for: {SearchTerm}", searchTerm);

                var allScopes = await GetAllScopesAsync();
                var lowerSearchTerm = searchTerm.ToLowerInvariant();

                var results = allScopes
                    .Where(scope => scope.ToLowerInvariant().Contains(lowerSearchTerm))
                    .Take(maxResults)
                    .ToList();

                _logger.Debug("Exact search completed: {ResultCount} results for {SearchTerm}", 
                    results.Count, searchTerm);

                return results;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Exact search failed for: {SearchTerm}", searchTerm);
                return new List<string>();
            }
        }

        private double CalculateFuzzyScore(string scope, string searchTerm, string[] searchWords)
        {
            try
            {
                var scopeLower = scope.ToLowerInvariant();
                var searchTermLower = searchTerm.ToLowerInvariant();
                double score = 0.0;

                // Extract or cache scope parts
                var (code, description) = GetScopeParts(scope);
                var codeLower = code.ToLowerInvariant();
                var descriptionLower = description.ToLowerInvariant();

                // Exact match bonus
                if (scopeLower == searchTermLower)
                {
                    score += 1000.0;
                }
                else if (scopeLower.Contains(searchTermLower))
                {
                    score += 500.0;
                }

                // Code match bonus (higher priority)
                if (codeLower.Contains(searchTermLower))
                {
                    score += 300.0;
                }
                else if (codeLower.StartsWith(searchTermLower))
                {
                    score += 200.0;
                }

                // Description match
                if (descriptionLower.Contains(searchTermLower))
                {
                    score += 100.0;
                }

                // Individual word matches
                foreach (var word in searchWords)
                {
                    if (word.Length < 2) continue; // Skip very short words

                    if (codeLower.Contains(word))
                    {
                        score += 50.0;
                    }
                    if (descriptionLower.Contains(word))
                    {
                        score += 25.0;
                    }
                }

                // Length penalty for very long descriptions
                if (description.Length > 50)
                {
                    score *= 0.9;
                }

                return score;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to calculate fuzzy score for scope: {Scope}", scope);
                return 0.0;
            }
        }

        private (string code, string description) GetScopeParts(string scope)
        {
            // Use cache to avoid repeated parsing
            if (_scopePartsCache.TryGetValue(scope, out var cached))
            {
                return cached;
            }

            try
            {
                // Parse scope format: "CODE - DESCRIPTION"
                var dashIndex = scope.IndexOf(" - ");
                if (dashIndex > 0)
                {
                    var code = scope.Substring(0, dashIndex).Trim();
                    var description = scope.Substring(dashIndex + 3).Trim();
                    var result = (code, description);
                    
                    // Cache for future use
                    _scopePartsCache[scope] = result;
                    return result;
                }
                else
                {
                    // Fallback if format is unexpected
                    var result = (scope, scope);
                    _scopePartsCache[scope] = result;
                    return result;
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to parse scope parts for: {Scope}", scope);
                var fallback = (scope, scope);
                _scopePartsCache[scope] = fallback;
                return fallback;
            }
        }
    }
} 