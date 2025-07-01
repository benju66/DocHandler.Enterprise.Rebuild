using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class FileNamePortionService
    {
        private readonly ILogger _logger;
        private readonly string _dataPath;
        private readonly string _portionsPath;
        private readonly string _recentPortionsPath;
        private FileNamePortionsData _data;
        private RecentPortionsData _recentData;
        
        public List<FileNamePortion> Portions => _data.Portions;
        public List<string> RecentPortions => _recentData.RecentPortions;
        
        public FileNamePortionService()
        {
            _logger = Log.ForContext<FileNamePortionService>();
            
            // Store data in AppData
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "DocHandler");
            Directory.CreateDirectory(appFolder);
            
            _dataPath = appFolder;
            _portionsPath = Path.Combine(appFolder, "filename_portions.json");
            _recentPortionsPath = Path.Combine(appFolder, "recent_portions.json");
            
            _data = LoadFileNamePortions();
            _recentData = LoadRecentPortions();
        }
        
        private FileNamePortionsData LoadFileNamePortions()
        {
            try
            {
                if (File.Exists(_portionsPath))
                {
                    var json = File.ReadAllText(_portionsPath);
                    var data = JsonSerializer.Deserialize<FileNamePortionsData>(json);
                    
                    if (data != null)
                    {
                        _logger.Information("Loaded {Count} filename portions", data.Portions.Count);
                        return data;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load filename portions");
            }
            
            // Return default data with construction scope codes
            _logger.Information("Creating default filename portions data");
            return CreateDefaultData();
        }
        
        private FileNamePortionsData CreateDefaultData()
        {
            return new FileNamePortionsData
            {
                Portions = new List<FileNamePortion>
                {
                    new FileNamePortion { Code = "02-4100", Description = "Demolition", Category = "Site Work" },
                    new FileNamePortion { Code = "03-1000", Description = "Concrete", Category = "Concrete" },
                    new FileNamePortion { Code = "04-2000", Description = "Masonry", Category = "Masonry" },
                    new FileNamePortion { Code = "05-1200", Description = "Structural Steel", Category = "Metals" },
                    new FileNamePortion { Code = "06-1000", Description = "Rough Carpentry", Category = "Wood & Plastics" },
                    new FileNamePortion { Code = "07-5000", Description = "Roofing", Category = "Thermal & Moisture" },
                    new FileNamePortion { Code = "08-1000", Description = "Doors and Frames", Category = "Doors & Windows" },
                    new FileNamePortion { Code = "09-2000", Description = "Gypsum Board", Category = "Finishes" },
                    new FileNamePortion { Code = "09-5000", Description = "Ceilings", Category = "Finishes" },
                    new FileNamePortion { Code = "09-6000", Description = "Flooring", Category = "Finishes" },
                    new FileNamePortion { Code = "15-0000", Description = "Mechanical", Category = "Mechanical" },
                    new FileNamePortion { Code = "16-0000", Description = "Electrical", Category = "Electrical" }
                }
            };
        }
        
        private RecentPortionsData LoadRecentPortions()
        {
            try
            {
                if (File.Exists(_recentPortionsPath))
                {
                    var json = File.ReadAllText(_recentPortionsPath);
                    var data = JsonSerializer.Deserialize<RecentPortionsData>(json);
                    
                    if (data != null)
                    {
                        _logger.Information("Loaded {Count} recent portions", data.RecentPortions.Count);
                        return data;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load recent portions");
            }
            
            return new RecentPortionsData { RecentPortions = new List<string>() };
        }
        
        public async Task SaveFileNamePortions()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_data, options);
                await File.WriteAllTextAsync(_portionsPath, json);
                
                _logger.Information("Filename portions saved to {Path}", _portionsPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save filename portions");
            }
        }
        
        public async Task SaveRecentPortions()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_recentData, options);
                await File.WriteAllTextAsync(_recentPortionsPath, json);
                
                _logger.Information("Recent portions saved to {Path}", _recentPortionsPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save recent portions");
            }
        }
        
        public async Task<bool> AddFileNamePortion(string code, string description, string category = "")
        {
            try
            {
                code = code.Trim();
                description = description.Trim();
                
                // Check if already exists
                if (_data.Portions.Any(p => p.Code.Equals(code, StringComparison.OrdinalIgnoreCase)))
                {
                    _logger.Warning("Filename portion already exists: {Code}", code);
                    return false;
                }
                
                var portion = new FileNamePortion
                {
                    Code = code,
                    Description = description,
                    Category = category,
                    DateAdded = DateTime.Now,
                    UsageCount = 0
                };
                
                _data.Portions.Add(portion);
                _data.Portions = _data.Portions.OrderBy(p => p.Code).ToList();
                
                await SaveFileNamePortions();
                _logger.Information("Added new filename portion: {Code} - {Description}", code, description);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to add filename portion: {Code}", code);
                return false;
            }
        }
        
        public async Task<bool> RemoveFileNamePortion(string code)
        {
            try
            {
                var portion = _data.Portions.FirstOrDefault(p => 
                    p.Code.Equals(code, StringComparison.OrdinalIgnoreCase));
                
                if (portion != null)
                {
                    _data.Portions.Remove(portion);
                    await SaveFileNamePortions();
                    _logger.Information("Removed filename portion: {Code}", code);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to remove filename portion: {Code}", code);
                return false;
            }
        }
        
        public async Task UpdateRecentPortion(string portionText)
        {
            try
            {
                // Remove if already exists
                _recentData.RecentPortions.Remove(portionText);
                
                // Add to beginning
                _recentData.RecentPortions.Insert(0, portionText);
                
                // Keep only max 20 recent portions
                while (_recentData.RecentPortions.Count > 20)
                {
                    _recentData.RecentPortions.RemoveAt(_recentData.RecentPortions.Count - 1);
                }
                
                await SaveRecentPortions();
                
                // Also update usage count in main data
                var portion = _data.Portions.FirstOrDefault(p => 
                    $"{p.Code} - {p.Description}".Equals(portionText, StringComparison.OrdinalIgnoreCase));
                
                if (portion != null)
                {
                    portion.UsageCount++;
                    portion.LastUsed = DateTime.Now;
                    await SaveFileNamePortions();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to update recent portion: {Portion}", portionText);
            }
        }
        
        public async Task ClearRecentPortions()
        {
            _recentData.RecentPortions.Clear();
            await SaveRecentPortions();
            _logger.Information("Cleared recent portions");
        }
        
        public string GetFormattedPortion(FileNamePortion portion)
        {
            return $"{portion.Code} - {portion.Description}";
        }
        
        public List<FileNamePortion> SearchPortions(string searchTerm)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
                return _data.Portions;
            
            searchTerm = searchTerm.ToLowerInvariant();
            
            return _data.Portions
                .Where(p => p.Code.ToLowerInvariant().Contains(searchTerm) ||
                           p.Description.ToLowerInvariant().Contains(searchTerm) ||
                           p.Category.ToLowerInvariant().Contains(searchTerm))
                .OrderBy(p => p.Code)
                .ToList();
        }
        
        public List<FileNamePortion> GetMostUsedPortions(int count = 10)
        {
            return _data.Portions
                .OrderByDescending(p => p.UsageCount)
                .ThenByDescending(p => p.LastUsed)
                .Take(count)
                .ToList();
        }
        
        public List<string> GetCategories()
        {
            return _data.Portions
                .Where(p => !string.IsNullOrWhiteSpace(p.Category))
                .Select(p => p.Category)
                .Distinct()
                .OrderBy(c => c)
                .ToList();
        }
    }
    
    public class FileNamePortionsData
    {
        public List<FileNamePortion> Portions { get; set; } = new();
    }
    
    public class FileNamePortion
    {
        public string Code { get; set; } = "";
        public string Description { get; set; } = "";
        public string Category { get; set; } = "";
        public DateTime DateAdded { get; set; } = DateTime.Now;
        public DateTime? LastUsed { get; set; }
        public int UsageCount { get; set; }
    }
    
    public class RecentPortionsData
    {
        public List<string> RecentPortions { get; set; } = new();
        public int MaxRecentPortions { get; set; } = 20;
    }
}