using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Microsoft.Extensions.DependencyInjection;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Models;
using DocHandler.Services;
using Serilog;

namespace DocHandler.ViewModels
{
    public partial class MainViewModelAsync : ObservableObject, IDisposable
    {
        private readonly ILogger _logger;
        private readonly IServiceProvider _serviceProvider;
        private readonly IConfigurationService _configService;
        private readonly CompanyNameService _companyService;
        private readonly ScopeOfWorkService _scopeService;
        private readonly IProcessManager _processManager;
        private readonly PdfCacheService _pdfCacheService;
        private readonly PerformanceMonitor _performanceMonitor;
        private readonly IOfficeServiceFactory _officeServiceFactory;
        
        // Office services - initialized asynchronously
        private ISessionAwareOfficeService? _sessionOfficeService;
        private ISessionAwareExcelService? _sessionExcelService;
        private IOfficeConversionService? _officeConversionService;
        
        // Service state
        private bool _isInitialized = false;
        private bool _disposed = false;

        [ObservableProperty]
        private ObservableCollection<FileItem> _droppedFiles = new();

        [ObservableProperty]
        private string _selectedCompany = "";

        [ObservableProperty]
        private string _selectedScope = "";

        [ObservableProperty]
        private string _saveLocation = "";

        [ObservableProperty]
        private bool _isProcessing = false;

        [ObservableProperty]
        private string _statusMessage = "Ready";

        [ObservableProperty]
        private bool _officeAvailable = false;

        public ObservableCollection<string> CompanyNames { get; }
        public ObservableCollection<string> ScopeOptions { get; }
        public ObservableCollection<string> RecentCompanies { get; }
        public ObservableCollection<string> RecentScopes { get; }

        // Commands
        public ICommand RemoveFileCommand { get; }
        public ICommand ClearAllFilesCommand { get; }
        public ICommand SelectSaveLocationCommand { get; }
        public ICommand ProcessFilesCommand { get; }
        public ICommand OpenSettingsCommand { get; }
        public ICommand OpenScopeEditorCommand { get; }
        public ICommand OpenCompanyEditorCommand { get; }

        public MainViewModelAsync(IServiceProvider serviceProvider)
        {
            _logger = Log.ForContext<MainViewModelAsync>();
            _serviceProvider = serviceProvider;

            // Get services from DI container using interfaces where available
            _configService = serviceProvider.GetRequiredService<IConfigurationService>();
            _companyService = serviceProvider.GetRequiredService<CompanyNameService>();
            _scopeService = serviceProvider.GetRequiredService<ScopeOfWorkService>();
            _processManager = serviceProvider.GetRequiredService<IProcessManager>();
            _pdfCacheService = serviceProvider.GetRequiredService<PdfCacheService>();
            _performanceMonitor = serviceProvider.GetRequiredService<PerformanceMonitor>();
            _officeServiceFactory = serviceProvider.GetRequiredService<IOfficeServiceFactory>();

            // Initialize collections
            CompanyNames = new ObservableCollection<string>();
            ScopeOptions = new ObservableCollection<string>(_scopeService.Scopes.Select(s => s.Code));
            RecentCompanies = new ObservableCollection<string>();
            RecentScopes = new ObservableCollection<string>(_scopeService.RecentScopes);

            // Initialize commands
            RemoveFileCommand = new RelayCommand<FileItem>(RemoveFile);
            ClearAllFilesCommand = new RelayCommand(ClearAllFiles);
            SelectSaveLocationCommand = new RelayCommand(SelectSaveLocation);
            ProcessFilesCommand = new AsyncRelayCommand(ProcessFilesAsync);
            OpenSettingsCommand = new RelayCommand(OpenSettings);
            OpenScopeEditorCommand = new RelayCommand(OpenScopeEditor);
            OpenCompanyEditorCommand = new RelayCommand(OpenCompanyEditor);

            // Load initial settings
            LoadSettings();

            _logger.Information("MainViewModelAsync created successfully");

            // Initialize Office services asynchronously
            _ = InitializeOfficeServicesAsync();
        }

        private async Task InitializeOfficeServicesAsync()
        {
            try
            {
                _logger.Information("Starting async Office service initialization...");
                
                // Check if Office is available first
                var isAvailable = await _officeServiceFactory.IsOfficeAvailableAsync();
                OfficeAvailable = isAvailable;
                
                if (!isAvailable)
                {
                    _logger.Warning("Microsoft Office is not available - running in limited mode");
                    StatusMessage = "Ready (Office unavailable)";
                    return;
                }

                // Initialize Office services with timeout protection
                var initializationTimeout = TimeSpan.FromSeconds(30);
                
                var officeTask = _officeServiceFactory.CreateOfficeServiceAsync();
                var sessionOfficeTask = _officeServiceFactory.CreateSessionOfficeServiceAsync();
                var sessionExcelTask = _officeServiceFactory.CreateSessionExcelServiceAsync();

                // Wait for all services with timeout
                var timeoutTask = Task.Delay(initializationTimeout);
                var allTasks = Task.WhenAll(officeTask, sessionOfficeTask, sessionExcelTask);
                
                if (await Task.WhenAny(allTasks, timeoutTask) == timeoutTask)
                {
                    _logger.Warning("Office service initialization timed out after {Timeout} seconds", initializationTimeout.TotalSeconds);
                    StatusMessage = "Ready (Office initialization timed out)";
                    return;
                }

                // Get the results
                _officeConversionService = await officeTask;
                _sessionOfficeService = await sessionOfficeTask;
                _sessionExcelService = await sessionExcelTask;

                // Set Office services for company name detection
                _companyService.SetOfficeServices(
                    _sessionOfficeService as SessionAwareOfficeService,
                    _sessionExcelService as SessionAwareExcelService);

                _isInitialized = true;
                StatusMessage = "Ready";
                
                _logger.Information("Office services initialized successfully");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to initialize Office services - continuing without Office support");
                StatusMessage = "Ready (Office services unavailable)";
                OfficeAvailable = false;
            }
        }

        private void LoadSettings()
        {
            try
            {
                // TODO: Load settings when AppConfiguration properties are available
                _logger.Debug("Settings loading skipped - configuration properties not available");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to load settings");
            }
        }

        private void SaveSettings()
        {
            try
            {
                // TODO: Save settings when AppConfiguration properties are available
                _logger.Debug("Settings saving skipped - configuration properties not available");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to save settings");
            }
        }

        public void HandleFileDrop(string[] files)
        {
            _performanceMonitor.StartOperation("FileDrop");
            
            try
            {
                foreach (var file in files)
                {
                    if (File.Exists(file) && IsFileSupported(file))
                    {
                        var fileItem = new FileItem
                        {
                            FilePath = file,
                            FileName = Path.GetFileName(file)
                        };

                        // Check for duplicates
                        if (!DroppedFiles.Any(f => f.FilePath.Equals(file, StringComparison.OrdinalIgnoreCase)))
                        {
                            DroppedFiles.Add(fileItem);
                            _logger.Information("Added file: {FileName}", fileItem.FileName);
                        }
                    }
                    else
                    {
                        _logger.Warning("Unsupported or invalid file: {File}", file);
                    }
                }

                UpdateStatus();
            }
            finally
            {
                _performanceMonitor.EndOperation("FileDrop");
            }
        }

        private bool IsFileSupported(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLower();
            return extension is ".pdf" or ".doc" or ".docx" or ".xls" or ".xlsx";
        }

        private void RemoveFile(FileItem? file)
        {
            if (file != null && DroppedFiles.Contains(file))
            {
                DroppedFiles.Remove(file);
                UpdateStatus();
                _logger.Information("Removed file: {FileName}", file.FileName);
            }
        }

        private void ClearAllFiles()
        {
            var count = DroppedFiles.Count;
            DroppedFiles.Clear();
            UpdateStatus();
            _logger.Information("Cleared {Count} files", count);
        }

        private void SelectSaveLocation()
        {
            try
            {
                using var dialog = new System.Windows.Forms.FolderBrowserDialog()
                {
                    Description = "Select save location for processed files",
                    ShowNewFolderButton = true,
                    SelectedPath = SaveLocation
                };

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    SaveLocation = dialog.SelectedPath;
                    SaveSettings();
                    _logger.Information("Save location updated: {Location}", SaveLocation);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to select save location");
                MessageBox.Show($"Error selecting folder: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ProcessFilesAsync()
        {
            if (IsProcessing) return;

            try
            {
                // Validate inputs
                if (!DroppedFiles.Any())
                {
                    MessageBox.Show("Please add files to process.", "No Files", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (string.IsNullOrEmpty(SaveLocation))
                {
                    MessageBox.Show("Please select a save location.", "No Save Location", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (string.IsNullOrEmpty(SelectedCompany))
                {
                    MessageBox.Show("Please select a company.", "No Company Selected", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                if (string.IsNullOrEmpty(SelectedScope))
                {
                    MessageBox.Show("Please select a scope of work.", "No Scope Selected", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                IsProcessing = true;
                StatusMessage = "Processing files...";
                
                _performanceMonitor.StartOperation("ProcessFiles");
                
                await Task.Run(() => ProcessFilesInternal());
                
                StatusMessage = "Processing completed successfully";
                _logger.Information("File processing completed successfully");
                
                // Save settings after successful processing
                SaveSettings();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during file processing");
                StatusMessage = "Processing failed";
                MessageBox.Show($"Error processing files: {ex.Message}", "Processing Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;
                _performanceMonitor.EndOperation("ProcessFiles");
            }
        }

        private void ProcessFilesInternal()
        {
            // This would contain the actual file processing logic
            // For now, just simulate processing
            foreach (var file in DroppedFiles)
            {
                Application.Current.Dispatcher.Invoke(() => 
                {
                    StatusMessage = $"Processing {file.FileName}...";
                });
                
                // Simulate processing time
                Task.Delay(500).Wait();
                
                _logger.Information("Processed file: {FileName}", file.FileName);
            }
        }

        private void UpdateStatus()
        {
            var fileCount = DroppedFiles.Count;
            if (fileCount == 0)
            {
                StatusMessage = "Ready - Drop files here";
            }
            else
            {
                StatusMessage = $"{fileCount} file{(fileCount == 1 ? "" : "s")} ready for processing";
            }
        }

        private void OpenSettings()
        {
            _logger.Information("Opening settings window");
            // TODO: Implement settings window
        }

        private void OpenScopeEditor()
        {
            _logger.Information("Opening scope editor");
            // TODO: Implement scope editor
        }

        private void OpenCompanyEditor()
        {
            _logger.Information("Opening company editor");
            // TODO: Implement company editor
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                try
                {
                    SaveSettings();
                    
                    _sessionOfficeService?.Dispose();
                    _sessionExcelService?.Dispose();
                    _officeConversionService?.Dispose();
                    
                    _logger.Information("MainViewModelAsync disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error during disposal");
                }
                
                _disposed = true;
            }
        }
    }
} 