using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using Serilog;
using FolderBrowserDialog = Ookii.Dialogs.Wpf.VistaFolderBrowserDialog;

namespace DocHandler.ViewModels
{
    public partial class SettingsViewModel : ObservableObject
    {
        private readonly ILogger _logger;
        private readonly ConfigurationService _configService;
        private readonly CompanyNameService _companyNameService;
        private readonly ScopeOfWorkService _scopeOfWorkService;
        private bool _hasUnsavedChanges = false;
        
        // Track original values for cancel functionality
        private AppConfiguration _originalConfig;
        
        public SettingsViewModel(
            ConfigurationService configService,
            CompanyNameService companyNameService,
            ScopeOfWorkService scopeOfWorkService)
        {
            _logger = Log.ForContext<SettingsViewModel>();
            _configService = configService;
            _companyNameService = companyNameService;
            _scopeOfWorkService = scopeOfWorkService;
            
            // Clone current config for cancel functionality
            _originalConfig = CloneConfiguration(_configService.Config);
            
            // Initialize collections
            LogLevels = new ObservableCollection<string> { "Debug", "Information", "Warning", "Error" };
            Themes = new ObservableCollection<string> { "Light", "Dark", "System" };
            
            LoadCurrentSettings();
        }
        
        #region General Settings
        
        [ObservableProperty]
        private string _defaultSaveLocation = "";
        
        [ObservableProperty]
        private string _selectedTheme = "Light";
        
        [ObservableProperty]
        private bool _rememberWindowPosition = true;
        
        [ObservableProperty]
        private bool _openFolderAfterProcessing = true;
        
        [ObservableProperty]
        private int _maxRecentLocations = 10;
        
        #endregion
        
        #region Processing Settings
        
        [ObservableProperty]
        private bool _saveQuotesMode = true;
        
        [ObservableProperty]
        private bool _convertOfficeToPdf = true;
        
        [ObservableProperty]
        private bool _autoScanCompanyNames = true;
        
        [ObservableProperty]
        private bool _scanCompanyNamesForDocFiles = false;
        
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(DocFileSizeLimitText))]
        private int _docFileSizeLimitMB = 10;
        
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(ParallelProcessingText))]
        private int _maxParallelProcessing = 3;
        
        [ObservableProperty]
        private int _conversionTimeoutSeconds = 30;
        
        [ObservableProperty]
        private bool _clearScopeAfterProcessing = false;
        
        #endregion
        
        #region Display Settings
        
        [ObservableProperty]
        private bool _showRecentScopes = false;
        
        [ObservableProperty]
        private bool _restoreQueueWindowOnStartup = true;
        
        [ObservableProperty]
        private bool _enableAnimations = true;
        
        [ObservableProperty]
        private bool _showStatusNotifications = true;
        
        #endregion
        
        #region Performance Settings
        
        [ObservableProperty]
        private int _memoryUsageLimitMB = 500;
        
        [ObservableProperty]
        private bool _enablePdfCaching = true;
        
        [ObservableProperty]
        private int _pdfCacheExpirationMinutes = 30;
        
        [ObservableProperty]
        private bool _enableProgressReporting = true;
        
        [ObservableProperty]
        private bool _cleanupTempFilesOnExit = true;
        
        #endregion
        
        #region Advanced Settings
        
        [ObservableProperty]
        private string _selectedLogLevel = "Information";
        
        [ObservableProperty]
        private string _logFileLocation = "";
        
        [ObservableProperty]
        private bool _enableDiagnosticMode = false;
        
        [ObservableProperty]
        private int _comTimeoutSeconds = 30;
        
        [ObservableProperty]
        private bool _enableNetworkPathOptimization = true;
        
        #endregion
        
        // Collections for dropdowns
        public ObservableCollection<string> LogLevels { get; }
        public ObservableCollection<string> Themes { get; }
        
        // Display text for sliders
        public string DocFileSizeLimitText => $"{DocFileSizeLimitMB} MB";
        public string ParallelProcessingText => $"{MaxParallelProcessing} concurrent files";
        
        // Commands
        [RelayCommand]
        private void BrowseDefaultSaveLocation()
        {
            var dialog = new FolderBrowserDialog
            {
                Description = "Select Default Save Location",
                SelectedPath = DefaultSaveLocation
            };
            
            if (dialog.ShowDialog() == true)
            {
                DefaultSaveLocation = dialog.SelectedPath;
                MarkAsChanged();
            }
        }
        
        [RelayCommand]
        private void BrowseLogFileLocation()
        {
            var dialog = new FolderBrowserDialog
            {
                Description = "Select Log File Location",
                SelectedPath = LogFileLocation
            };
            
            if (dialog.ShowDialog() == true)
            {
                LogFileLocation = dialog.SelectedPath;
                MarkAsChanged();
            }
        }
        
        [RelayCommand]
        private void EditCompanyNames()
        {
            // Open company names editor
            var editor = new Views.EditCompanyNamesWindow(_companyNameService)
            {
                Owner = Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive)
            };
            editor.ShowDialog();
        }
        
        [RelayCommand]
        private void EditScopesOfWork()
        {
            // Open scopes editor
            var editor = new Views.EditScopesOfWorkWindow(_scopeOfWorkService)
            {
                Owner = Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive)
            };
            editor.ShowDialog();
        }
        
        [RelayCommand]
        private void ClearCache()
        {
            var result = MessageBox.Show(
                "This will clear all cached PDF files. Continue?",
                "Clear Cache",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
                
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    // Clear PDF cache
                    var cacheDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "DocHandler", "PdfCache");
                        
                    if (Directory.Exists(cacheDir))
                    {
                        Directory.Delete(cacheDir, true);
                        Directory.CreateDirectory(cacheDir);
                    }
                    
                    MessageBox.Show("Cache cleared successfully.", "Success",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                        
                    _logger.Information("Cache cleared by user");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to clear cache");
                    MessageBox.Show($"Failed to clear cache: {ex.Message}", "Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        
        [RelayCommand]
        private void ResetToDefaults()
        {
            var result = MessageBox.Show(
                "This will reset all settings to their default values. Continue?",
                "Reset Settings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
                
            if (result == MessageBoxResult.Yes)
            {
                // Reset to default configuration
                var defaults = _configService.GetDefaultConfiguration();
                LoadFromConfiguration(defaults);
                MarkAsChanged();
                
                _logger.Information("Settings reset to defaults");
            }
        }
        
        [RelayCommand]
        private void ExportSettings()
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                FileName = $"DocHandler_Settings_{DateTime.Now:yyyyMMdd}.json"
            };
            
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    SaveToConfiguration();
                    var json = System.Text.Json.JsonSerializer.Serialize(_configService.Config, 
                        new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                    File.WriteAllText(dialog.FileName, json);
                    
                    MessageBox.Show("Settings exported successfully.", "Export Complete",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                        
                    _logger.Information("Settings exported to {Path}", dialog.FileName);
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to export settings");
                    MessageBox.Show($"Failed to export settings: {ex.Message}", "Export Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        
        [RelayCommand]
        private void ImportSettings()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
            };
            
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var json = File.ReadAllText(dialog.FileName);
                    var imported = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(json);
                    
                    if (imported != null)
                    {
                        LoadFromConfiguration(imported);
                        MarkAsChanged();
                        
                        MessageBox.Show("Settings imported successfully.", "Import Complete",
                            MessageBoxButton.OK, MessageBoxImage.Information);
                            
                        _logger.Information("Settings imported from {Path}", dialog.FileName);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to import settings");
                    MessageBox.Show($"Failed to import settings: {ex.Message}", "Import Error",
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        
        public bool SaveSettings()
        {
            try
            {
                SaveToConfiguration();
                _ = _configService.SaveConfiguration();
                
                _hasUnsavedChanges = false;
                _originalConfig = CloneConfiguration(_configService.Config);
                
                _logger.Information("Settings saved successfully");
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save settings");
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Save Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }
        
        public void CancelChanges()
        {
            LoadFromConfiguration(_originalConfig);
            _hasUnsavedChanges = false;
        }
        
        private void LoadCurrentSettings()
        {
            LoadFromConfiguration(_configService.Config);
        }
        
        private void LoadFromConfiguration(AppConfiguration config)
        {
            // General
            DefaultSaveLocation = config.DefaultSaveLocation;
            SelectedTheme = config.Theme;
            RememberWindowPosition = config.RememberWindowPosition;
            OpenFolderAfterProcessing = config.OpenFolderAfterProcessing ?? true;
            MaxRecentLocations = config.MaxRecentLocations;
            
            // Processing
            SaveQuotesMode = config.SaveQuotesMode;
            AutoScanCompanyNames = config.AutoScanCompanyNames;
            ScanCompanyNamesForDocFiles = config.ScanCompanyNamesForDocFiles;
            DocFileSizeLimitMB = config.DocFileSizeLimitMB;
            MaxParallelProcessing = config.MaxParallelProcessing;
            ConversionTimeoutSeconds = config.ConversionTimeoutSeconds;
            ClearScopeAfterProcessing = config.ClearScopeAfterProcessing;
            
            // Display
            ShowRecentScopes = config.ShowRecentScopes;
            RestoreQueueWindowOnStartup = config.RestoreQueueWindowOnStartup;
            EnableAnimations = config.EnableAnimations;
            ShowStatusNotifications = config.ShowStatusNotifications;
            
            // Performance
            MemoryUsageLimitMB = config.MemoryUsageLimitMB;
            EnablePdfCaching = config.EnablePdfCaching;
            PdfCacheExpirationMinutes = config.PdfCacheExpirationMinutes;
            EnableProgressReporting = config.EnableProgressReporting;
            CleanupTempFilesOnExit = config.CleanupTempFilesOnExit;
            
            // Advanced
            SelectedLogLevel = config.LogLevel;
            LogFileLocation = config.LogFileLocation;
            EnableDiagnosticMode = config.EnableDiagnosticMode;
            ComTimeoutSeconds = config.ComTimeoutSeconds;
            EnableNetworkPathOptimization = config.EnableNetworkPathOptimization;
        }
        
        private void SaveToConfiguration()
        {
            var config = _configService.Config;
            
            // General
            config.DefaultSaveLocation = DefaultSaveLocation;
            config.Theme = SelectedTheme;
            config.RememberWindowPosition = RememberWindowPosition;
            config.OpenFolderAfterProcessing = OpenFolderAfterProcessing;
            config.MaxRecentLocations = MaxRecentLocations;
            
            // Processing
            config.SaveQuotesMode = SaveQuotesMode;
            config.AutoScanCompanyNames = AutoScanCompanyNames;
            config.ScanCompanyNamesForDocFiles = ScanCompanyNamesForDocFiles;
            config.DocFileSizeLimitMB = DocFileSizeLimitMB;
            config.MaxParallelProcessing = MaxParallelProcessing;
            config.ConversionTimeoutSeconds = ConversionTimeoutSeconds;
            config.ClearScopeAfterProcessing = ClearScopeAfterProcessing;
            
            // Display
            config.ShowRecentScopes = ShowRecentScopes;
            config.RestoreQueueWindowOnStartup = RestoreQueueWindowOnStartup;
            config.EnableAnimations = EnableAnimations;
            config.ShowStatusNotifications = ShowStatusNotifications;
            
            // Performance
            config.MemoryUsageLimitMB = MemoryUsageLimitMB;
            config.EnablePdfCaching = EnablePdfCaching;
            config.PdfCacheExpirationMinutes = PdfCacheExpirationMinutes;
            config.EnableProgressReporting = EnableProgressReporting;
            config.CleanupTempFilesOnExit = CleanupTempFilesOnExit;
            
            // Advanced
            config.LogLevel = SelectedLogLevel;
            config.LogFileLocation = LogFileLocation;
            config.EnableDiagnosticMode = EnableDiagnosticMode;
            config.ComTimeoutSeconds = ComTimeoutSeconds;
            config.EnableNetworkPathOptimization = EnableNetworkPathOptimization;
            
            // Update services with new settings
            _companyNameService.UpdateDocFileSizeLimit(DocFileSizeLimitMB);
        }
        
        private AppConfiguration CloneConfiguration(AppConfiguration config)
        {
            // Create a deep copy of the configuration
            var json = System.Text.Json.JsonSerializer.Serialize(config);
            return System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(json)!;
        }
        
        private void MarkAsChanged()
        {
            _hasUnsavedChanges = true;
        }
        
        protected override void OnPropertyChanged(PropertyChangedEventArgs e)
        {
            base.OnPropertyChanged(e);
            
            // Mark as changed for any property except display texts
            if (e.PropertyName != nameof(DocFileSizeLimitText) && 
                e.PropertyName != nameof(ParallelProcessingText))
            {
                MarkAsChanged();
            }
        }
        
        public bool HasUnsavedChanges => _hasUnsavedChanges;
    }
} 