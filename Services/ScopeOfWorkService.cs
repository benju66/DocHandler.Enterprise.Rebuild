using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class ScopeOfWorkService
    {
        private readonly ILogger _logger;
        private readonly string _dataPath;
        private readonly string _scopesPath;
        private readonly string _recentScopesPath;
        private readonly string _defaultScopesPath;
        private ScopeOfWorkData _data;
        private RecentScopesData _recentData;
        
        public List<ScopeOfWork> Scopes => _data.Scopes;
        public List<string> RecentScopes => _recentData.RecentScopes;
        
        public ScopeOfWorkService()
        {
            _logger = Log.ForContext<ScopeOfWorkService>();
            
            // Store data in AppData
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "DocHandler");
            Directory.CreateDirectory(appFolder);
            
            _dataPath = appFolder;
            _scopesPath = Path.Combine(appFolder, "scopes_of_work.json");
            _recentScopesPath = Path.Combine(appFolder, "recent_scopes.json");
            _defaultScopesPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "default_scopes.json");
            
            _data = LoadScopesOfWork();
            _recentData = LoadRecentScopes();
        }
        
        private ScopeOfWorkData LoadScopesOfWork()
        {
            try
            {
                // First check if user has their own scopes file
                if (File.Exists(_scopesPath))
                {
                    var json = File.ReadAllText(_scopesPath);
                    var data = JsonSerializer.Deserialize<ScopeOfWorkData>(json);
                    
                    if (data != null && data.Scopes != null && data.Scopes.Any())
                    {
                        _logger.Information("Loaded {Count} scopes of work from user data", data.Scopes.Count);
                        return data;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load user scopes of work");
            }
            
            // If no user data, load defaults
            return LoadDefaultScopes();
        }
        
        private ScopeOfWorkData LoadDefaultScopes()
        {
            try
            {
                // Try to load from default scopes file
                if (File.Exists(_defaultScopesPath))
                {
                    var json = File.ReadAllText(_defaultScopesPath);
                    var data = JsonSerializer.Deserialize<ScopeOfWorkData>(json);
                    
                    if (data != null && data.Scopes != null && data.Scopes.Any())
                    {
                        // Initialize runtime properties for each scope
                        foreach (var scope in data.Scopes)
                        {
                            if (scope.DateAdded == default(DateTime))
                                scope.DateAdded = DateTime.Now;
                            scope.UsageCount = 0;
                            scope.LastUsed = null;
                        }
                        
                        _logger.Information("Loaded {Count} default scopes from file", data.Scopes.Count);
                        
                        // Save to user's location for future editing
                        _ = SaveScopesOfWork();
                        
                        return data;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Could not load default scopes file from {Path}", _defaultScopesPath);
            }
            
            // Return minimal starter set if file not found
            _logger.Warning("Default scopes file not found, creating minimal set");
            return new ScopeOfWorkData
            {
                Scopes = new List<ScopeOfWork>
                {
                    new ScopeOfWork { Code = "00-0000", Description = "General", DateAdded = DateTime.Now }
                }
            };
        }
        
        private RecentScopesData LoadRecentScopes()
        {
            try
            {
                if (File.Exists(_recentScopesPath))
                {
                    var json = File.ReadAllText(_recentScopesPath);
                    var data = JsonSerializer.Deserialize<RecentScopesData>(json);
                    
                    if (data != null)
                    {
                        _logger.Information("Loaded {Count} recent scopes", data.RecentScopes.Count);
                        return data;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load recent scopes");
            }
            
            return new RecentScopesData { RecentScopes = new List<string>() };
        }
        
        public async Task SaveScopesOfWork()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_data, options);
                await File.WriteAllTextAsync(_scopesPath, json);
                
                _logger.Information("Scopes of work saved to {Path}", _scopesPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save scopes of work");
                throw;
            }
        }
        
        public async Task SaveRecentScopes()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_recentData, options);
                await File.WriteAllTextAsync(_recentScopesPath, json);
                
                _logger.Information("Recent scopes saved to {Path}", _recentScopesPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save recent scopes");
            }
        }
        
        public async Task<bool> AddScope(string code, string description)
        {
            try
            {
                code = code.Trim();
                description = description.Trim();
                
                // Validate
                if (string.IsNullOrWhiteSpace(code) || string.IsNullOrWhiteSpace(description))
                {
                    _logger.Warning("Cannot add scope with empty code or description");
                    return false;
                }
                
                // Check if already exists
                if (_data.Scopes.Any(s => s.Code.Equals(code, StringComparison.OrdinalIgnoreCase)))
                {
                    _logger.Warning("Scope of work already exists: {Code}", code);
                    return false;
                }
                
                var scope = new ScopeOfWork
                {
                    Code = code,
                    Description = description,
                    DateAdded = DateTime.Now,
                    UsageCount = 0
                };
                
                _data.Scopes.Add(scope);
                _data.Scopes = _data.Scopes.OrderBy(s => s.Code).ToList();
                
                await SaveScopesOfWork();
                _logger.Information("Added new scope of work: {Code} - {Description}", code, description);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to add scope of work: {Code}", code);
                return false;
            }
        }
        
        public async Task<bool> UpdateScope(string oldCode, string newCode, string newDescription)
        {
            try
            {
                var scope = _data.Scopes.FirstOrDefault(s => 
                    s.Code.Equals(oldCode, StringComparison.OrdinalIgnoreCase));
                
                if (scope == null)
                {
                    _logger.Warning("Scope not found for update: {Code}", oldCode);
                    return false;
                }
                
                newCode = newCode.Trim();
                newDescription = newDescription.Trim();
                
                // Validate
                if (string.IsNullOrWhiteSpace(newCode) || string.IsNullOrWhiteSpace(newDescription))
                {
                    _logger.Warning("Cannot update scope with empty code or description");
                    return false;
                }
                
                // Check if new code conflicts with existing (unless it's the same scope)
                if (!oldCode.Equals(newCode, StringComparison.OrdinalIgnoreCase) &&
                    _data.Scopes.Any(s => s.Code.Equals(newCode, StringComparison.OrdinalIgnoreCase)))
                {
                    _logger.Warning("Cannot update scope: code {Code} already exists", newCode);
                    return false;
                }
                
                scope.Code = newCode;
                scope.Description = newDescription;
                
                // Resort the list
                _data.Scopes = _data.Scopes.OrderBy(s => s.Code).ToList();
                
                await SaveScopesOfWork();
                _logger.Information("Updated scope: {OldCode} -> {NewCode} - {Description}", 
                    oldCode, newCode, newDescription);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to update scope: {Code}", oldCode);
                return false;
            }
        }
        
        public async Task<bool> RemoveScope(string code)
        {
            try
            {
                var scope = _data.Scopes.FirstOrDefault(s => 
                    s.Code.Equals(code, StringComparison.OrdinalIgnoreCase));
                
                if (scope != null)
                {
                    _data.Scopes.Remove(scope);
                    await SaveScopesOfWork();
                    _logger.Information("Removed scope of work: {Code}", code);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to remove scope of work: {Code}", code);
                return false;
            }
        }
        
        public async Task UpdateRecentScope(string scopeText)
        {
            try
            {
                // Remove if already exists
                _recentData.RecentScopes.Remove(scopeText);
                
                // Add to beginning
                _recentData.RecentScopes.Insert(0, scopeText);
                
                // Keep only max recent scopes
                while (_recentData.RecentScopes.Count > _recentData.MaxRecentScopes)
                {
                    _recentData.RecentScopes.RemoveAt(_recentData.RecentScopes.Count - 1);
                }
                
                await SaveRecentScopes();
                
                // Also update usage count in main data
                var scope = _data.Scopes.FirstOrDefault(s => 
                    GetFormattedScope(s).Equals(scopeText, StringComparison.OrdinalIgnoreCase));
                
                if (scope != null)
                {
                    scope.UsageCount++;
                    scope.LastUsed = DateTime.Now;
                    await SaveScopesOfWork();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to update recent scope: {Scope}", scopeText);
            }
        }
        
        public async Task ClearRecentScopes()
        {
            _recentData.RecentScopes.Clear();
            await SaveRecentScopes();
            _logger.Information("Cleared recent scopes");
        }
        
        public string GetFormattedScope(ScopeOfWork scope)
        {
            return $"{scope.Code} - {scope.Description}";
        }
        
        public List<ScopeOfWork> SearchScopes(string searchTerm)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
                return _data.Scopes;
            
            searchTerm = searchTerm.ToLowerInvariant();
            
            return _data.Scopes
                .Where(s => s.Code.ToLowerInvariant().Contains(searchTerm) ||
                           s.Description.ToLowerInvariant().Contains(searchTerm))
                .OrderBy(s => s.Code)
                .ToList();
        }
        
        public List<ScopeOfWork> GetMostUsedScopes(int count = 10)
        {
            return _data.Scopes
                .Where(s => s.UsageCount > 0)
                .OrderByDescending(s => s.UsageCount)
                .ThenByDescending(s => s.LastUsed)
                .Take(count)
                .ToList();
        }
        
        public async Task<ImportResult> ImportScopes(string filePath, bool replace = false)
        {
            var result = new ImportResult();
            
            try
            {
                if (!File.Exists(filePath))
                {
                    result.Success = false;
                    result.Message = "File not found";
                    return result;
                }
                
                var json = await File.ReadAllTextAsync(filePath);
                var importData = JsonSerializer.Deserialize<ScopeOfWorkData>(json);
                
                if (importData == null || importData.Scopes == null || !importData.Scopes.Any())
                {
                    result.Success = false;
                    result.Message = "No valid scopes found in file";
                    return result;
                }
                
                if (replace)
                {
                    _data.Scopes.Clear();
                }
                
                foreach (var importScope in importData.Scopes)
                {
                    if (!_data.Scopes.Any(s => s.Code.Equals(importScope.Code, StringComparison.OrdinalIgnoreCase)))
                    {
                        _data.Scopes.Add(new ScopeOfWork
                        {
                            Code = importScope.Code,
                            Description = importScope.Description,
                            DateAdded = importScope.DateAdded == default ? DateTime.Now : importScope.DateAdded,
                            UsageCount = 0,
                            LastUsed = null
                        });
                        result.Added++;
                    }
                    else
                    {
                        result.Skipped++;
                    }
                    result.TotalProcessed++;
                }
                
                // Resort and save
                _data.Scopes = _data.Scopes.OrderBy(s => s.Code).ToList();
                await SaveScopesOfWork();
                
                result.Success = true;
                result.Message = $"Import completed: {result.Added} added, {result.Skipped} skipped";
                _logger.Information(result.Message);
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to import scopes from {Path}", filePath);
                result.Success = false;
                result.Message = $"Import failed: {ex.Message}";
                return result;
            }
        }
        
        public async Task<bool> ExportScopes(string filePath)
        {
            try
            {
                var exportData = new
                {
                    Scopes = _data.Scopes.Select(s => new
                    {
                        s.Code,
                        s.Description
                    })
                };
                
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(exportData, options);
                await File.WriteAllTextAsync(filePath, json);
                
                _logger.Information("Exported {Count} scopes to {Path}", _data.Scopes.Count, filePath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to export scopes to {Path}", filePath);
                return false;
            }
        }
        
        public async Task<bool> ResetToDefaults()
        {
            try
            {
                _data = LoadDefaultScopes();
                await SaveScopesOfWork();
                _logger.Information("Reset scopes to defaults");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to reset scopes to defaults");
                return false;
            }
        }
    }
    
    public class ScopeOfWorkData
    {
        public List<ScopeOfWork> Scopes { get; set; } = new();
    }
    
    public class ScopeOfWork
    {
        public string Code { get; set; } = "";
        public string Description { get; set; } = "";
        public DateTime DateAdded { get; set; } = DateTime.Now;
        public DateTime? LastUsed { get; set; }
        public int UsageCount { get; set; }
    }
    
    public class RecentScopesData
    {
        public List<string> RecentScopes { get; set; } = new();
        public int MaxRecentScopes { get; set; } = 20;
    }
    
    public class ImportResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public int TotalProcessed { get; set; }
        public int Added { get; set; }
        public int Skipped { get; set; }
    }
}