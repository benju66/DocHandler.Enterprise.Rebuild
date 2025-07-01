using System;
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
        
        public ConfigurationService ConfigService => _configService;
        
        [ObservableProperty]
        private ObservableCollection<FileItem> _pendingFiles = new();

        // Checkbox options - hardcoded for now
        private bool _convertOfficeToPdf = true;
        public bool ConvertOfficeToPdf 
        { 
            get => _convertOfficeToPdf;
            set => SetProperty(ref _convertOfficeToPdf, value);
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
                        StatusMessage = "Save Quotes Mode: Select a scope of work and drop quotes";
                        SessionSaveLocation = _configService.Config.DefaultSaveLocation;
                    }
                    else
                    {
                        StatusMessage = "Drop files here to begin";
                        SelectedScope = null;
                    }
                }
            }
        }

        [ObservableProperty]
        private string? _selectedScope;

        [ObservableProperty]
        private ObservableCollection<string> _scopesOfWork = new();

        [ObservableProperty]
        private ObservableCollection<string> _recentScopes = new();

        [ObservableProperty]
        private string _scopeSearchText = "";

        [ObservableProperty]
        private string _sessionSaveLocation = "";
        
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
            
            // Initialize theme from config
            IsDarkMode = _configService.Config.Theme == "Dark";
            
            // Update UI when files are added/removed
            PendingFiles.CollectionChanged += (s, e) => UpdateUI();
            
            // Check Office availability
            CheckOfficeAvailability();
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
        }

        private void LoadRecentScopes()
        {
            RecentScopes.Clear();
            foreach (var scope in _scopeOfWorkService.RecentScopes.Take(10))
            {
                RecentScopes.Add(scope);
            }
        }
        
        private void UpdateUI()
        {
            if (SaveQuotesMode)
            {
                CanProcess = PendingFiles.Count > 0 && !IsProcessing && !string.IsNullOrEmpty(SelectedScope);
                ProcessButtonText = PendingFiles.Count > 1 ? "Process All Quotes" : "Process Quote";
                
                if (PendingFiles.Count == 0)
                {
                    StatusMessage = string.IsNullOrEmpty(SelectedScope) 
                        ? "Select a scope of work and drop quotes" 
                        : $"Selected: {SelectedScope} - Drop quote documents";
                }
                else
                {
                    StatusMessage = $"{PendingFiles.Count} quote(s) ready - {SelectedScope}";
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

                    // Update configuration with recent location
                    _configService.AddRecentLocation(outputDir);

                    // Open the output folder
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
            StatusMessage = "Files cleared";
            UpdateUI();
        }
        
        [RelayCommand]
        private void RemoveFile(FileItem? fileItem)
        {
            if (fileItem != null)
            {
                PendingFiles.Remove(fileItem);
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
                StatusMessage = $"Selected: {scope} - Drop quote documents";
            }
            UpdateUI();
        }

        [RelayCommand]
        private void SearchScopes()
        {
            ScopesOfWork.Clear();
            var searchResults = _scopeOfWorkService.SearchScopes(ScopeSearchText);
            
            foreach (var scope in searchResults)
            {
                ScopesOfWork.Add(_scopeOfWorkService.GetFormattedScope(scope));
            }
        }

        [RelayCommand]
        private async Task ClearRecentScopes()
        {
            await _scopeOfWorkService.ClearRecentScopes();
            LoadRecentScopes();
        }

        private async Task ProcessSaveQuotes()
        {
            if (!SaveQuotesMode || string.IsNullOrEmpty(SelectedScope))
            {
                MessageBox.Show("Please select a scope of work first.", "Save Quotes", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
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
                        
                        // First, scan for company name
                        var detectedCompany = await _companyNameService.ScanDocumentForCompanyName(file.FilePath);
                        string companyName;

                        if (!string.IsNullOrEmpty(detectedCompany))
                        {
                            // Found a company - confirm with user
                            var confirmResult = MessageBox.Show(
                                $"Detected company: {detectedCompany}\n\nIs this correct?",
                                "Confirm Company Name",
                                MessageBoxButton.YesNoCancel,
                                MessageBoxImage.Question);

                            if (confirmResult == MessageBoxResult.Cancel)
                                continue;

                            if (confirmResult == MessageBoxResult.Yes)
                            {
                                companyName = detectedCompany;
                            }
                            else
                            {
                                // User said no, prompt for correct name
                                companyName = await PromptForCompanyName();
                                if (string.IsNullOrEmpty(companyName))
                                    continue;
                            }
                        }
                        else
                        {
                            // No company detected, prompt user
                            companyName = await PromptForCompanyName();
                            if (string.IsNullOrEmpty(companyName))
                                continue;
                        }

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

                // Update status
                if (processedCount > 0)
                {
                    StatusMessage = $"Successfully processed {processedCount} quote(s)";
                    
                    // Open output folder
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

        private async Task<string> PromptForCompanyName(string detectedCompany = null)
        {
            return await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                string message = detectedCompany != null
                    ? $"Company '{detectedCompany}' was detected. Confirm or enter a different company name:"
                    : "Please enter the company name for this quote:";

                // Create suggested companies collection
                // TODO: Update this when you add a GetCompanies/GetCompanyNames method to CompanyNameService
                var suggestedCompanies = new ObservableCollection<string>();

                // Create and show the dialog
                var dialog = new CompanyNameDialog(message, suggestedCompanies)
                {
                    CompanyName = detectedCompany ?? string.Empty,
                    Owner = Application.Current.MainWindow
                };

                if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.CompanyName))
                {
                    string companyName = dialog.CompanyName.Trim();

                    // Add to database if requested
                    if (dialog.AddToDatabase)
                    {
                        _ = _companyNameService.AddCompanyName(companyName);
                    }

                    return companyName;
                }

                return null;
            });
        }
        
        // Placeholder commands - implement these later
        [RelayCommand]
        private void SetSaveLocation()
        {
            // TODO: Implement folder browser dialog
            MessageBox.Show("Set Save Location - Coming Soon", "Feature", 
                MessageBoxButton.OK, MessageBoxImage.Information);
        }

        [RelayCommand]
        private void EditCompanyNames()
        {
            // TODO: Implement company names editor window
            MessageBox.Show("Edit Company Names - Coming Soon", "Feature", 
                MessageBoxButton.OK, MessageBoxImage.Information);
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
        
        public void Cleanup()
        {
            _officeConversionService?.Dispose();
        }
        
        public void SaveWindowState(double left, double top, double width, double height, string state)
        {
            if (_configService.Config.RememberWindowPosition)
            {
                _configService.UpdateWindowPosition(left, top, width, height, state);
                _ = _configService.SaveConfiguration();
            }
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