using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using DocHandler.Views;
using Serilog;
using MessageBox = System.Windows.MessageBox;
using Application = System.Windows.Application;
using FolderBrowserDialog = Ookii.Dialogs.Wpf.VistaFolderBrowserDialog;

namespace DocHandler.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger _logger;
        private readonly FileProcessingService _fileProcessingService;
        private readonly ConfigurationService _configService;
        private readonly OfficeConversionService _officeConversionService;
        private readonly CompanyNameService _companyNameService;
        private readonly ScopeOfWorkService _scopeOfWorkService;
        private readonly List<string> _tempFilesToCleanup = new();
        
        public ConfigurationService ConfigService => _configService;
        
        [ObservableProperty]
        private ObservableCollection<FileItem> _pendingFiles = new();

        // Checkbox options
        private bool _convertOfficeToPdf = true;
        public bool ConvertOfficeToPdf 
        { 
            get => _convertOfficeToPdf;
            set => SetProperty(ref _convertOfficeToPdf, value);
        }

        private bool _openFolderAfterProcessing = true;
        public bool OpenFolderAfterProcessing
        {
            get => _openFolderAfterProcessing;
            set => SetProperty(ref _openFolderAfterProcessing, value);
        }
        
        [ObservableProperty]
        private bool _isProcessing;
        
        [ObservableProperty]
        private double _progressValue;
        
        [ObservableProperty]
        private string _statusMessage = "Drop files here to begin";
        
        [ObservableProperty]
        private bool _canProcess;
        
        [ObservableProperty]
        private string _processButtonText = "Process Files";
        
        // Save Quotes Mode properties
        private bool _saveQuotesMode;
        public bool SaveQuotesMode
        {
            get => _saveQuotesMode;
            set
            {
                if (SetProperty(ref _saveQuotesMode, value))
                {
                    UpdateUI();
                    
                    if (value)
                    {
                        StatusMessage = "Save Quotes Mode: Drop quote documents";
                        SessionSaveLocation = _configService.Config.DefaultSaveLocation;
                    }
                    else
                    {
                        StatusMessage = "Drop files here to begin";
                        SelectedScope = null;
                        CompanyNameInput = "";
                        DetectedCompanyName = "";
                    }
                }
            }
        }

        [ObservableProperty]
        private string? _selectedScope;

        [ObservableProperty]
        private ObservableCollection<string> _scopesOfWork = new();

        [ObservableProperty]
        private ObservableCollection<string> _filteredScopesOfWork = new();

        [ObservableProperty]
        private ObservableCollection<string> _recentScopes = new();

        private string _scopeSearchText = "";
        public string ScopeSearchText
        {
            get => _scopeSearchText;
            set
            {
                if (SetProperty(ref _scopeSearchText, value))
                {
                    FilterScopes();
                }
            }
        }

        [ObservableProperty]
        private string _sessionSaveLocation = "";

        // Company name fields
        [ObservableProperty]
        private string _companyNameInput = "";

        [ObservableProperty]
        private string _detectedCompanyName = "";

        [ObservableProperty]
        private bool _isDetectingCompany;

        // Recent locations
        public ObservableCollection<string> RecentLocations => 
            new ObservableCollection<string>(_configService.Config.RecentLocations);
        
        public MainViewModel()
        {
            _logger = Log.ForContext<MainViewModel>();
            _fileProcessingService = new FileProcessingService();
            _configService = new ConfigurationService();
            _officeConversionService = new OfficeConversionService();
            _companyNameService = new CompanyNameService();
            _scopeOfWorkService = new ScopeOfWorkService();
            
            // Load scopes of work
            LoadScopesOfWork();
            LoadRecentScopes();
            
            // Initialize session save location
            SessionSaveLocation = _configService.Config.DefaultSaveLocation;
            
            // Initialize theme from config
            IsDarkMode = _configService.Config.Theme == "Dark";
            
            // Load open folder preference
            OpenFolderAfterProcessing = _configService.Config.OpenFolderAfterProcessing ?? true;
            
            // Update UI when files are added/removed
            PendingFiles.CollectionChanged += (s, e) => 
            {
                UpdateUI();
                
                // When files are added in Save Quotes mode, scan for company names
                if (SaveQuotesMode && e.NewItems != null)
                {
                    foreach (FileItem item in e.NewItems)
                    {
                        _ = ScanForCompanyName(item.FilePath);
                    }
                }
            };
            
            // Check Office availability
            CheckOfficeAvailability();
        }
        
        private async Task ScanForCompanyName(string filePath)
        {
            if (!SaveQuotesMode || IsDetectingCompany) return;
            
            try
            {
                IsDetectingCompany = true;
                _logger.Information("Scanning document for company name: {Path}", filePath);
                
                var detectedCompany = await _companyNameService.ScanDocumentForCompanyName(filePath);
                
                if (!string.IsNullOrEmpty(detectedCompany))
                {
                    // Only update if user hasn't typed anything
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = detectedCompany;
                        _logger.Information("Detected company name: {Company}", detectedCompany);
                        
                        // Force UI update after detection
                        UpdateUI();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to scan document for company name");
            }
            finally
            {
                IsDetectingCompany = false;
                UpdateUI(); // Ensure UI updates after detection completes
            }
        }
        
        private void CheckOfficeAvailability()
        {
            if (!_officeConversionService.IsOfficeInstalled())
            {
                _logger.Warning("Microsoft Office is not available - Word/Excel conversion features will be disabled");
            }
        }
        
        private void LoadScopesOfWork()
        {
            ScopesOfWork.Clear();
            foreach (var scope in _scopeOfWorkService.Scopes)
            {
                ScopesOfWork.Add(_scopeOfWorkService.GetFormattedScope(scope));
            }
            FilterScopes();
        }

        private void LoadRecentScopes()
        {
            RecentScopes.Clear();
            foreach (var scope in _scopeOfWorkService.RecentScopes.Take(10))
            {
                RecentScopes.Add(scope);
            }
        }

        private void FilterScopes()
        {
            FilteredScopesOfWork.Clear();
            
            var searchTerm = ScopeSearchText?.Trim() ?? "";
            var filteredScopes = string.IsNullOrWhiteSpace(searchTerm) 
                ? ScopesOfWork 
                : ScopesOfWork.Where(s => s.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0);
            
            foreach (var scope in filteredScopes)
            {
                FilteredScopesOfWork.Add(scope);
            }
        }
        
        public void UpdateUI()
        {
            if (SaveQuotesMode)
            {
                // Need both a scope and either typed or detected company name
                var hasCompanyName = !string.IsNullOrWhiteSpace(CompanyNameInput) || 
                                   !string.IsNullOrWhiteSpace(DetectedCompanyName);
                
                CanProcess = PendingFiles.Count > 0 && !IsProcessing && 
                           !string.IsNullOrEmpty(SelectedScope) && hasCompanyName;
                
                ProcessButtonText = PendingFiles.Count > 1 ? "Process All Quotes" : "Process Quote";
                
                if (PendingFiles.Count == 0)
                {
                    StatusMessage = "Save Quotes Mode: Drop quote documents";
                }
                else
                {
                    StatusMessage = $"{PendingFiles.Count} quote(s) ready to process";
                }
            }
            else
            {
                CanProcess = PendingFiles.Count > 0 && !IsProcessing;
                ProcessButtonText = PendingFiles.Count > 1 ? "Merge and Save" : "Process Files";
                
                if (PendingFiles.Count == 0)
                {
                    StatusMessage = "Drop files here to begin";
                }
                else
                {
                    StatusMessage = $"{PendingFiles.Count} file(s) ready to process";
                }
            }
        }
        
        public void AddFiles(string[] filePaths)
        {
            var validFiles = _fileProcessingService.ValidateDroppedFiles(filePaths);
            
            foreach (var file in validFiles)
            {
                // Check if file already added
                if (PendingFiles.Any(f => f.FilePath == file))
                {
                    _logger.Information("File already in list: {FilePath}", file);
                    continue;
                }
                
                var fileItem = new FileItem
                {
                    FilePath = file,
                    FileName = Path.GetFileName(file),
                    FileSize = new FileInfo(file).Length,
                    FileType = Path.GetExtension(file).ToUpperInvariant().TrimStart('.')
                };
                
                PendingFiles.Add(fileItem);
            }
            
            if (validFiles.Count != filePaths.Length)
            {
                var invalidCount = filePaths.Length - validFiles.Count;
                MessageBox.Show($"{invalidCount} file(s) were not added because they are not supported.", 
                    "Some Files Not Added", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        /// <summary>
        /// Adds temporary files that should be cleaned up after processing
        /// </summary>
        public void AddTempFilesForCleanup(List<string> tempFiles)
        {
            _tempFilesToCleanup.AddRange(tempFiles);
            _logger.Debug("Added {Count} temp files for cleanup", tempFiles.Count);
        }
        
        private void CleanupTempFiles()
        {
            foreach (var tempFile in _tempFilesToCleanup)
            {
                try
                {
                    if (File.Exists(tempFile))
                    {
                        File.Delete(tempFile);
                        _logger.Debug("Cleaned up temp file: {File}", tempFile);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to cleanup temp file: {File}", tempFile);
                }
            }
            _tempFilesToCleanup.Clear();
        }
        
        [RelayCommand]
        private async Task ProcessFiles()
        {
            if (SaveQuotesMode)
            {
                await ProcessSaveQuotes();
                return;
            }

            if (!PendingFiles.Any())
            {
                StatusMessage = "No files selected";
                return;
            }

            IsProcessing = true;
            StatusMessage = PendingFiles.Count > 1 ? "Merging and processing files..." : "Processing file...";

            try
            {
                var filePaths = PendingFiles.Select(f => f.FilePath).ToList();
                var outputDir = _configService.Config.DefaultSaveLocation;

                // Create output folder with timestamp
                outputDir = _fileProcessingService.CreateOutputFolder(outputDir);

                var result = await _fileProcessingService.ProcessFiles(filePaths, outputDir, ConvertOfficeToPdf);

                if (result.Success)
                {
                    if (result.IsMerged)
                    {
                        StatusMessage = $"Successfully merged {filePaths.Count} files into {Path.GetFileName(result.SuccessfulFiles.First())}";
                        _logger.Information("Files merged successfully");
                    }
                    else
                    {
                        StatusMessage = $"Successfully processed {result.SuccessfulFiles.Count} file(s)";
                        _logger.Information("Files processed successfully");
                    }

                    // Clear the file list after successful processing
                    PendingFiles.Clear();
                    
                    // Clean up any temp files
                    CleanupTempFiles();

                    // Update configuration with recent location
                    _configService.AddRecentLocation(outputDir);

                    // Open the output folder if preference is set
                    if (OpenFolderAfterProcessing)
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = outputDir,
                                UseShellExecute = true,
                                Verb = "open"
                            });
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Failed to open output folder");
                        }
                    }
                }
                else
                {
                    var errorMessage = !string.IsNullOrEmpty(result.ErrorMessage) 
                        ? result.ErrorMessage 
                        : "Processing failed";
                    
                    StatusMessage = $"Error: {errorMessage}";
                    _logger.Error("File processing failed: {Error}", errorMessage);

                    if (result.FailedFiles.Any())
                    {
                        var failedFilesList = string.Join("\n", result.FailedFiles.Select(f => 
                            $"• {Path.GetFileName(f.FilePath)}: {f.Error}"));
                        
                        MessageBox.Show(
                            $"The following files could not be processed:\n\n{failedFilesList}",
                            "Processing Errors",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                _logger.Error(ex, "Unexpected error during file processing");
                MessageBox.Show(
                    $"An unexpected error occurred:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;
                ProgressValue = 0;
                UpdateUI();
            }
        }
        
        [RelayCommand]
        private void ClearFiles()
        {
            PendingFiles.Clear();
            CleanupTempFiles();
            
            // Reset company name detection
            CompanyNameInput = "";
            DetectedCompanyName = "";
            
            StatusMessage = "Files cleared";
            UpdateUI();
        }
        
        [RelayCommand]
        private void RemoveFile(FileItem? fileItem)
        {
            if (fileItem != null)
            {
                PendingFiles.Remove(fileItem);
                
                // If this was a temp file, clean it up immediately
                if (_tempFilesToCleanup.Contains(fileItem.FilePath))
                {
                    try
                    {
                        if (File.Exists(fileItem.FilePath))
                        {
                            File.Delete(fileItem.FilePath);
                            _logger.Debug("Cleaned up removed temp file: {File}", fileItem.FilePath);
                        }
                        _tempFilesToCleanup.Remove(fileItem.FilePath);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to cleanup removed temp file: {File}", fileItem.FilePath);
                    }
                }
                
                // If no files left, clear company detection
                if (PendingFiles.Count == 0 && SaveQuotesMode)
                {
                    CompanyNameInput = "";
                    DetectedCompanyName = "";
                }
                
                UpdateUI();
            }
        }

        [RelayCommand]
        public void SelectScope(string? scope)
        {
            SelectedScope = scope;
            if (!string.IsNullOrEmpty(scope))
            {
                _ = _scopeOfWorkService.UpdateRecentScope(scope);
                LoadRecentScopes();
            }
            UpdateUI();
        }

        [RelayCommand]
        private void SearchScopes()
        {
            FilterScopes();
        }

        [RelayCommand]
        private async Task ClearRecentScopes()
        {
            await _scopeOfWorkService.ClearRecentScopes();
            LoadRecentScopes();
        }

        [RelayCommand]
        private void SetSaveLocation()
        {
            var dialog = new FolderBrowserDialog
            {
                Description = "Select save location for documents",
                UseDescriptionForTitle = true
            };
            
            // Set initial directory
            if (Directory.Exists(SessionSaveLocation))
            {
                dialog.SelectedPath = SessionSaveLocation;
            }
            else if (Directory.Exists(_configService.Config.DefaultSaveLocation))
            {
                dialog.SelectedPath = _configService.Config.DefaultSaveLocation;
            }
            
            if (dialog.ShowDialog() == true)
            {
                SessionSaveLocation = dialog.SelectedPath;
                _configService.UpdateDefaultSaveLocation(dialog.SelectedPath);
                _ = _configService.SaveConfiguration();
                
                // Notify UI of recent locations change
                OnPropertyChanged(nameof(RecentLocations));
                
                _logger.Information("Save location set to: {Path}", dialog.SelectedPath);
            }
        }

        [RelayCommand]
        private void OpenSaveLocation()
        {
            if (Directory.Exists(SessionSaveLocation))
            {
                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = SessionSaveLocation,
                        UseShellExecute = true,
                        Verb = "open"
                    });
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to open save location");
                    MessageBox.Show("Could not open the save location folder.", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            else
            {
                MessageBox.Show("The save location folder does not exist.", "Folder Not Found", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        [RelayCommand]
        private void ShowRecentLocations()
        {
            var dialog = new RecentLocationsDialog(_configService.Config.RecentLocations)
            {
                Owner = Application.Current.MainWindow
            };
            
            if (dialog.ShowDialog() == true && !string.IsNullOrEmpty(dialog.SelectedLocation))
            {
                SessionSaveLocation = dialog.SelectedLocation;
                _configService.UpdateDefaultSaveLocation(dialog.SelectedLocation);
                _ = _configService.SaveConfiguration();
                
                // Notify UI of recent locations change
                OnPropertyChanged(nameof(RecentLocations));
                
                _logger.Information("Save location set to: {Path}", dialog.SelectedLocation);
            }
        }

        [RelayCommand]
        private void EditCompanyNames()
        {
            var window = new EditCompanyNamesWindow(_companyNameService)
            {
                Owner = Application.Current.MainWindow
            };
            
            if (window.ShowDialog() == true)
            {
                _logger.Information("Company names were modified");
                
                // If we're in Save Quotes mode and have files, rescan them for company names
                // since the database has changed
                if (SaveQuotesMode && PendingFiles.Any())
                {
                    var firstFile = PendingFiles.First();
                    CompanyNameInput = "";
                    DetectedCompanyName = "";
                    _ = ScanForCompanyName(firstFile.FilePath);
                }
            }
        }

        [RelayCommand]
        private void EditScopesOfWork()
        {
            // TODO: Implement scopes of work editor window
            MessageBox.Show("Edit Scopes of Work - Coming Soon", "Feature", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [RelayCommand]
        private void Exit()
        {
            Application.Current.Shutdown();
        }

        [RelayCommand]
        private void About()
        {
            MessageBox.Show(
                "DocHandler Enterprise\nVersion 1.0\n\nDocument Processing Tool with Save Quotes Mode\n\n© 2024", 
                "About DocHandler",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        private async Task ProcessSaveQuotes()
        {
            if (!SaveQuotesMode || string.IsNullOrEmpty(SelectedScope))
            {
                MessageBox.Show("Please select a scope of work first.", "Save Quotes", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Get company name - use typed value first, then detected value
            var companyName = !string.IsNullOrWhiteSpace(CompanyNameInput) 
                ? CompanyNameInput.Trim() 
                : DetectedCompanyName?.Trim();
            
            if (string.IsNullOrWhiteSpace(companyName))
            {
                MessageBox.Show("Please enter a company name or wait for automatic detection.", 
                    "Company Name Required", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!PendingFiles.Any())
            {
                StatusMessage = "No quote documents to process";
                return;
            }

            IsProcessing = true;
            var processedCount = 0;
            var failedFiles = new List<(string file, string error)>();

            try
            {
                var outputDir = !string.IsNullOrEmpty(SessionSaveLocation) 
                    ? SessionSaveLocation 
                    : _configService.Config.DefaultSaveLocation;

                foreach (var file in PendingFiles.ToList())
                {
                    try
                    {
                        StatusMessage = $"Processing quote: {file.FileName}";

                        // Build the filename: [Scope] - [Company].pdf
                        var outputFileName = $"{SelectedScope} - {companyName}.pdf";
                        var outputPath = Path.Combine(outputDir, outputFileName);

                        // Ensure unique filename
                        outputPath = Path.Combine(outputDir, 
                            _fileProcessingService.GetUniqueFileName(outputDir, outputFileName));

                        // Process the file (convert if needed and save)
                        var processResult = await ProcessSingleQuoteFile(file.FilePath, outputPath);
                        
                        if (processResult.Success)
                        {
                            processedCount++;
                            PendingFiles.Remove(file);
                            _logger.Information("Saved quote as: {FileName}", outputFileName);
                            
                            // Update company usage if it was detected
                            if (!string.IsNullOrWhiteSpace(DetectedCompanyName) && 
                                companyName.Equals(DetectedCompanyName, StringComparison.OrdinalIgnoreCase))
                            {
                                await _companyNameService.IncrementUsageCount(companyName);
                            }
                            
                            // Add to company database if not already there
                            if (!_companyNameService.Companies.Any(c => 
                                c.Name.Equals(companyName, StringComparison.OrdinalIgnoreCase)))
                            {
                                await _companyNameService.AddCompanyName(companyName);
                            }
                        }
                        else
                        {
                            failedFiles.Add((file.FileName, processResult.ErrorMessage ?? "Unknown error"));
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to process quote: {File}", file.FileName);
                        failedFiles.Add((file.FileName, ex.Message));
                    }
                }

                // Clean up any temp files
                CleanupTempFiles();
                
                // Clear company name fields after processing
                CompanyNameInput = "";
                DetectedCompanyName = "";

                // Update status
                if (processedCount > 0)
                {
                    StatusMessage = $"Successfully processed {processedCount} quote(s)";
                    
                    // Update recent locations
                    _configService.AddRecentLocation(outputDir);
                    OnPropertyChanged(nameof(RecentLocations));
                    
                    // Open output folder if preference is set
                    if (OpenFolderAfterProcessing)
                    {
                        try
                        {
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = outputDir,
                                UseShellExecute = true,
                                Verb = "open"
                            });
                        }
                        catch { }
                    }
                }

                if (failedFiles.Any())
                {
                    var failedList = string.Join("\n", failedFiles.Select(f => $"• {f.file}: {f.error}"));
                    MessageBox.Show(
                        $"The following quotes could not be processed:\n\n{failedList}",
                        "Processing Errors",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Save Quotes processing failed");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsProcessing = false;
                UpdateUI();
            }
        }

        private async Task<ProcessingResult> ProcessSingleQuoteFile(string inputPath, string outputPath)
        {
            var extension = Path.GetExtension(inputPath).ToLowerInvariant();
            
            if (extension == ".pdf")
            {
                // Just copy the PDF
                File.Copy(inputPath, outputPath, true);
                return new ProcessingResult { Success = true, SuccessfulFiles = { outputPath } };
            }
            else
            {
                // Convert to PDF first
                var files = new List<string> { inputPath };
                var tempDir = _fileProcessingService.CreateTempFolder();
                
                try
                {
                    var result = await _fileProcessingService.ProcessFiles(files, tempDir, true);
                    
                    if (result.Success && result.SuccessfulFiles.Any())
                    {
                        var convertedPdf = result.SuccessfulFiles.First();
                        File.Move(convertedPdf, outputPath, true);
                        result.SuccessfulFiles[0] = outputPath;
                    }
                    
                    return result;
                }
                finally
                {
                    // Clean up temp folder
                    try { Directory.Delete(tempDir, true); } catch { }
                }
            }
        }
        
        // Command handlers
        partial void OnCompanyNameInputChanged(string value)
        {
            UpdateUI();
        }
        
        partial void OnDetectedCompanyNameChanged(string? value)
        {
            UpdateUI();
        }
        
        partial void OnSelectedScopeChanged(string? value)
        {
            UpdateUI();
        }
        
        public void Cleanup()
        {
            _officeConversionService?.Dispose();
            CleanupTempFiles();
        }
        
        public void SaveWindowState(double left, double top, double width, double height, string state)
        {
            if (_configService.Config.RememberWindowPosition)
            {
                _configService.UpdateWindowPosition(left, top, width, height, state);
                _ = _configService.SaveConfiguration();
            }
        }

        public void SavePreferences()
        {
            _configService.Config.OpenFolderAfterProcessing = OpenFolderAfterProcessing;
            _ = _configService.SaveConfiguration();
        }

        // Property for theme
        private bool _isDarkMode = false;
        public bool IsDarkMode
        {
            get => _isDarkMode;
            set
            {
                if (SetProperty(ref _isDarkMode, value))
                {
                    ModernWpf.ThemeManager.Current.ApplicationTheme = value
                        ? ModernWpf.ApplicationTheme.Dark
                        : ModernWpf.ApplicationTheme.Light;
                    _configService.UpdateTheme(value ? "Dark" : "Light");
                }
            }
        }
    }
    
    public class FileItem
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public string FileType { get; set; } = "";
        public long FileSize { get; set; }
        
        public string FileSizeDisplay => FormatFileSize(FileSize);
        
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