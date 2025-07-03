using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class ConfigurationService
    {
        private readonly ILogger _logger;
        private readonly string _configPath;
        private AppConfiguration _config;
        
        public AppConfiguration Config => _config;
        
        public ConfigurationService()
        {
            _logger = Log.ForContext<ConfigurationService>();
            
            // Store config in AppData
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "DocHandler");
            Directory.CreateDirectory(appFolder);
            
            _configPath = Path.Combine(appFolder, "config.json");
            _config = LoadConfiguration();
        }
        
        private AppConfiguration LoadConfiguration()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    var json = File.ReadAllText(_configPath);
                    var config = JsonSerializer.Deserialize<AppConfiguration>(json);
                    
                    if (config != null)
                    {
                        _logger.Information("Configuration loaded from {Path}", _configPath);
                        return config;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load configuration");
            }
            
            // Return default configuration
            _logger.Information("Using default configuration");
            return CreateDefaultConfiguration();
        }
        
        private AppConfiguration CreateDefaultConfiguration()
        {
            return new AppConfiguration
            {
                DefaultSaveLocation = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                RecentLocations = new List<string>(),
                MaxRecentLocations = 10,
                Theme = "Light",
                RememberWindowPosition = true,
                WindowLeft = 100,
                WindowTop = 100,
                WindowWidth = 800,
                WindowHeight = 600,
                WindowState = "Normal",
                OpenFolderAfterProcessing = true,
                SaveQuotesMode = true,  // Added - default to true
                ShowRecentScopes = false,  // Added - default to false (hidden by default)
                AutoScanCompanyNames = true  // Added - default to true (enabled)
            };
        }
        
        public async Task SaveConfiguration()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_config, options);
                await File.WriteAllTextAsync(_configPath, json);
                
                _logger.Information("Configuration saved to {Path}", _configPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save configuration");
            }
        }
        
        public void AddRecentLocation(string location)
        {
            if (string.IsNullOrWhiteSpace(location))
                return;
                
            // Remove if already exists
            _config.RecentLocations.Remove(location);
            
            // Add to beginning
            _config.RecentLocations.Insert(0, location);
            
            // Keep only max number of locations
            while (_config.RecentLocations.Count > _config.MaxRecentLocations)
            {
                _config.RecentLocations.RemoveAt(_config.RecentLocations.Count - 1);
            }
            
            // Save changes
            _ = SaveConfiguration();
        }
        
        public void UpdateWindowPosition(double left, double top, double width, double height, string state)
        {
            _config.WindowLeft = left;
            _config.WindowTop = top;
            _config.WindowWidth = width;
            _config.WindowHeight = height;
            _config.WindowState = state;
        }
        
        public void UpdateTheme(string theme)
        {
            _config.Theme = theme;
            _ = SaveConfiguration();
        }
        
        public void UpdateDefaultSaveLocation(string location)
        {
            _config.DefaultSaveLocation = location;
            AddRecentLocation(location);
            _ = SaveConfiguration();
        }
    }
    
    public class AppConfiguration
    {
        public string DefaultSaveLocation { get; set; } = "";
        public List<string> RecentLocations { get; set; } = new();
        public int MaxRecentLocations { get; set; } = 10;
        public string Theme { get; set; } = "Light";
        public bool RememberWindowPosition { get; set; } = true;
        public double WindowLeft { get; set; }
        public double WindowTop { get; set; }
        public double WindowWidth { get; set; }
        public double WindowHeight { get; set; }
        public string WindowState { get; set; } = "Normal";
        public bool? OpenFolderAfterProcessing { get; set; } = true;
        public bool SaveQuotesMode { get; set; } = true;  // Added - default to true
        public bool ShowRecentScopes { get; set; } = false;  // Added - default to false (hidden by default)
        public bool AutoScanCompanyNames { get; set; } = true;  // Added - default to true (enabled)
    }
}