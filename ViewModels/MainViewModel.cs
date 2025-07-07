// Folder: ViewModels/
// File: MainViewModel.cs
// Enhanced with Fuzzy Search for Scopes
// Fixed: Save Quotes Mode default and Scope synchronization
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using DocHandler.Views;
using Serilog;
using MessageBox = System.Windows.MessageBox;
using Application = System.Windows.Application;
using FolderBrowserDialog = Ookii.Dialogs.Wpf.VistaFolderBrowserDialog;
using System.Windows.Threading;

namespace DocHandler.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger _logger;
        private readonly OptimizedFileProcessingService _fileProcessingService;
        private readonly ConfigurationService _configService;
        private readonly OfficeConversionService _officeConversionService;
        private readonly CompanyNameService _companyNameService;
        private readonly ScopeOfWorkService _scopeOfWorkService;
        private readonly SessionAwareOfficeService _sessionOfficeService;
        private readonly object _conversionLock = new object();
        private readonly ConcurrentDictionary<string, byte> _tempFilesToCleanup = new();
        
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
                    
                    // Save the preference
                    _configService.Config.SaveQuotesMode = value;
                    _ = _configService.SaveConfiguration();
                    
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

        private string? _selectedScope;
        public string? SelectedScope
        {
            get => _selectedScope;
            set
            {
                if (SetProperty(ref _selectedScope, value))
                {
                    UpdateUI();
                }
            }
        }

        [ObservableProperty]
        private ObservableCollection<string> _scopesOfWork = new();

        [ObservableProperty]
        private ObservableCollection<string> _filteredScopesOfWork = new();

        [ObservableProperty]
        private ObservableCollection<string> _recentScopes = new();

        private string _scopeSearchText = "";
        private DispatcherTimer _scopeSearchTimer;
        private CancellationTokenSource _scopeSearchCancellation;
        private const int ScopeSearchDelayMs = 300;
        
        public string ScopeSearchText
        {
            get => _scopeSearchText;
            set
            {
                if (SetProperty(ref _scopeSearchText, value))
                {
                    // Cancel previous search
                    _scopeSearchCancellation?.Cancel();
                    _scopeSearchCancellation = new CancellationTokenSource();
                    
                    // Restart the timer for debouncing
                    _scopeSearchTimer?.Stop();
                    _scopeSearchTimer?.Start();
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

        // Performance tracking
        private int _processedFileCount = 0;

        // Recent locations
        public ObservableCollection<string> RecentLocations => 
            new ObservableCollection<string>(_configService.Config.RecentLocations);
        
        public MainViewModel()
        {
            _logger = Log.ForContext<MainViewModel>();
            _fileProcessingService = new OptimizedFileProcessingService();
            _configService = new ConfigurationService();
            _officeConversionService = new OfficeConversionService();
            _companyNameService = new CompanyNameService();
            _scopeOfWorkService = new ScopeOfWorkService();
            
            // Initialize session-aware Office service for better performance
            _sessionOfficeService = new SessionAwareOfficeService();
            _logger.Information("Session-aware Office service initialized");
            
            // Initialize scope search timer for debouncing
            _scopeSearchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(ScopeSearchDelayMs)
            };
            _scopeSearchTimer.Tick += (s, e) =>
            {
                _scopeSearchTimer.Stop();
                FilterScopes();
            };
            
            // Load Save Quotes Mode from config
            SaveQuotesMode = _configService.Config.SaveQuotesMode;
            
            // Load scopes of work
            LoadScopesOfWork();
            LoadRecentScopes();
            
            // Initialize session save location
            SessionSaveLocation = _configService.Config.DefaultSaveLocation;
            
            // Initialize theme from config
            IsDarkMode = _configService.Config.Theme == "Dark";
            
            // Load open folder preference
            OpenFolderAfterProcessing = _configService.Config.OpenFolderAfterProcessing ?? true;
            
            // Load show recent scopes preference
            ShowRecentScopes = _configService.Config.ShowRecentScopes;
            
            // Load auto-scan company names preference
            AutoScanCompanyNames = _configService.Config.AutoScanCompanyNames;
            
            // Update UI when files are added/removed
            PendingFiles.CollectionChanged += (s, e) => 
            {
                UpdateUI();
                
                // When files are added in Save Quotes mode, scan for company names
                // ONLY if user hasn't already entered a company name
                if (SaveQuotesMode && e.NewItems != null && string.IsNullOrWhiteSpace(CompanyNameInput))
                {
                    foreach (FileItem item in e.NewItems)
                    {
                        // Scan the first file added for company names
                        _ = ScanForCompanyName(item.FilePath);
                        break; // Only scan the first file to avoid multiple detections
                    }
                }
            };
            
            // Check Office availability
            CheckOfficeAvailability();
        }
        
        private async Task ScanForCompanyName(string filePath)
        {
            // Don't scan if auto-scan is disabled, user has already typed a company name, or other conditions
            if (!SaveQuotesMode || IsDetectingCompany || !string.IsNullOrWhiteSpace(CompanyNameInput) || !AutoScanCompanyNames) 
            {
                _logger.Debug("Skipping company name scan - SaveQuotesMode: {SaveQuotesMode}, IsDetecting: {IsDetecting}, HasInput: {HasInput}, AutoScan: {AutoScan}", 
                    SaveQuotesMode, IsDetectingCompany, !string.IsNullOrWhiteSpace(CompanyNameInput), AutoScanCompanyNames);
                return;
            }
            
            try
            {
                IsDetectingCompany = true;
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                _logger.Information("Starting optimized company name detection for: {Path}", filePath);
                
                // Use cancellation token for responsive UI with 30-second timeout
                using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                
                var detectedCompany = await Task.Run(async () =>
                {
                    return await _companyNameService.ScanDocumentForCompanyName(filePath);
                }, cts.Token);
                
                stopwatch.Stop();
                _logger.Information("Company detection completed in {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
                
                if (!string.IsNullOrEmpty(detectedCompany))
                {
                    // Only update if user still hasn't typed anything
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = detectedCompany;
                        _logger.Information("Successfully detected company: {Company} in {ElapsedMs}ms", 
                            detectedCompany, stopwatch.ElapsedMilliseconds);
                        
                        // Force UI update after detection
                        UpdateUI();
                    }
                    else
                    {
                        _logger.Information("Company detected ({Company}) but user has already typed: {UserInput}", 
                            detectedCompany, CompanyNameInput);
                    }
                }
                else
                {
                    _logger.Information("No company name detected in document: {Path} (took {ElapsedMs}ms)", 
                        filePath, stopwatch.ElapsedMilliseconds);
                    // Clear any previous detection
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = "";
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.Warning("Company detection timed out for: {Path}", filePath);
                // Show user-friendly message for timeout
                if (string.IsNullOrWhiteSpace(CompanyNameInput))
                {
                    DetectedCompanyName = "";
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Company detection failed for: {Path}", filePath);
                // Clear any partial detection on error
                if (string.IsNullOrWhiteSpace(CompanyNameInput))
                {
                    DetectedCompanyName = "";
                }
            }
            finally
            {
                IsDetectingCompany = false;
                _logger.Debug("Company name detection completed. DetectedCompanyName: '{DetectedName}'", 
                    DetectedCompanyName ?? "null");
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
            
            // Initial filter
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

        // Enhanced fuzzy search implementation with better synchronization
        private void FilterScopes()
        {
            var searchTerm = ScopeSearchText?.Trim() ?? "";
            
            // Get current selection before filtering
            var currentSelection = SelectedScope;
            
            // Build filtered list without clearing to minimize UI disruption
            var filteredList = new List<string>();
            
            if (string.IsNullOrWhiteSpace(searchTerm))
            {
                // No search term - show all scopes
                filteredList.AddRange(ScopesOfWork);
            }
            else
            {
                // Fuzzy search implementation
                var searchWords = searchTerm.ToLowerInvariant().Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
                
                var scoredScopes = new List<(string scope, double score)>();
                
                foreach (var scope in ScopesOfWork)
                {
                    double score = CalculateFuzzyScore(scope, searchTerm, searchWords);
                    if (score > 0)
                    {
                        scoredScopes.Add((scope, score));
                    }
                }
                
                // Sort by score (highest first) and add to filtered list
                filteredList.AddRange(scoredScopes.OrderByDescending(x => x.score).Select(x => x.scope));
            }
            
            // Update the filtered collection efficiently
            Application.Current.Dispatcher.Invoke(() =>
            {
                // Only update if the contents have changed
                if (!filteredList.SequenceEqual(FilteredScopesOfWork))
                {
                    FilteredScopesOfWork.Clear();
                    foreach (var scope in filteredList)
                    {
                        FilteredScopesOfWork.Add(scope);
                    }
                }
                
                // Preserve selection more conservatively
                if (currentSelection != null)
                {
                    if (FilteredScopesOfWork.Contains(currentSelection))
                    {
                        // Keep existing selection if it's still in filtered results
                        if (SelectedScope != currentSelection)
                        {
                            SelectedScope = currentSelection;
                        }
                    }
                    else if (string.IsNullOrWhiteSpace(searchTerm))
                    {
                        // Only clear selection if search is completely empty
                        // This prevents clearing during navigation
                        SelectedScope = null;
                    }
                    // Don't clear selection during active search - user might be navigating
                }
            });
        }
        
        private double CalculateFuzzyScore(string scope, string searchTerm, string[] searchWords)
        {
            var scopeLower = scope.ToLowerInvariant();
            var searchTermLower = searchTerm.ToLowerInvariant();
            
            // Split scope into code and description parts
            var dashIndex = scope.IndexOf(" - ");
            string code = dashIndex > 0 ? scope.Substring(0, dashIndex).ToLowerInvariant() : "";
            string description = dashIndex > 0 ? scope.Substring(dashIndex + 3).ToLowerInvariant() : scopeLower;
            
            double score = 0;
            
            // 1. Exact match (highest score)
            if (scopeLower == searchTermLower)
            {
                return 100;
            }
            
            // 2. Exact code match
            if (code == searchTermLower)
            {
                return 90;
            }
            
            // 3. Code starts with search term
            if (!string.IsNullOrEmpty(code) && code.StartsWith(searchTermLower))
            {
                score += 80 - (code.Length - searchTermLower.Length); // Closer matches score higher
            }
            
            // 4. Code contains search term
            else if (!string.IsNullOrEmpty(code) && code.Contains(searchTermLower))
            {
                score += 60;
            }
            
            // 5. Description exact match
            if (description == searchTermLower)
            {
                score += 85;
            }
            
            // 6. Description starts with search term
            else if (description.StartsWith(searchTermLower))
            {
                score += 70;
            }
            
            // 7. Full scope contains exact search term
            else if (scopeLower.Contains(searchTermLower))
            {
                score += 50;
                // Bonus if it's at a word boundary
                if (Regex.IsMatch(scopeLower, $@"\b{Regex.Escape(searchTermLower)}\b"))
                {
                    score += 10;
                }
            }
            
            // 8. All search words are found (in any order)
            if (searchWords.Length > 1)
            {
                bool allWordsFound = true;
                int wordsFoundCount = 0;
                
                foreach (var word in searchWords)
                {
                    if (scopeLower.Contains(word))
                    {
                        wordsFoundCount++;
                    }
                    else
                    {
                        allWordsFound = false;
                    }
                }
                
                if (allWordsFound)
                {
                    score += 40;
                }
                else if (wordsFoundCount > 0)
                {
                    // Partial match - score based on percentage of words found
                    score += 20 * ((double)wordsFoundCount / searchWords.Length);
                }
            }
            
            // 9. Individual word matching (for single word searches)
            else if (searchWords.Length == 1)
            {
                var searchWord = searchWords[0];
                
                // Check each word in the scope
                var scopeWords = scopeLower.Split(new[] { ' ', '-', ',', '.' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var scopeWord in scopeWords)
                {
                    if (scopeWord == searchWord)
                    {
                        score += 35; // Exact word match
                    }
                    else if (scopeWord.StartsWith(searchWord))
                    {
                        score += 25; // Word starts with search
                    }
                    else if (scopeWord.Contains(searchWord))
                    {
                        score += 15; // Word contains search
                    }
                }
            }
            
            // 10. Fuzzy matching for typos (Levenshtein distance)
            if (score == 0 && searchTermLower.Length >= 3) // Only for searches 3+ chars
            {
                // Check description words for close matches
                var descWords = description.Split(new[] { ' ', '-', ',', '.' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var word in descWords)
                {
                    var distance = LevenshteinDistance(searchTermLower, word);
                    var maxLen = Math.Max(searchTermLower.Length, word.Length);
                    var similarity = 1.0 - ((double)distance / maxLen);
                    
                    // If 80% similar or better, include it
                    if (similarity >= 0.8)
                    {
                        score += 10 * similarity;
                    }
                }
            }
            
            return score;
        }
        
        // Simple Levenshtein distance implementation for fuzzy matching
        private int LevenshteinDistance(string s1, string s2)
        {
            int[,] distance = new int[s1.Length + 1, s2.Length + 1];
            
            for (int i = 0; i <= s1.Length; i++)
                distance[i, 0] = i;
            
            for (int j = 0; j <= s2.Length; j++)
                distance[0, j] = j;
            
            for (int i = 1; i <= s1.Length; i++)
            {
                for (int j = 1; j <= s2.Length; j++)
                {
                    int cost = (s1[i - 1] == s2[j - 1]) ? 0 : 1;
                    distance[i, j] = Math.Min(
                        Math.Min(distance[i - 1, j] + 1, distance[i, j - 1] + 1),
                        distance[i - 1, j - 1] + cost
                    );
                }
            }
            
            return distance[s1.Length, s2.Length];
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
            foreach (var tempFile in tempFiles)
            {
                _tempFilesToCleanup.TryAdd(tempFile, 0);
            }
            _logger.Debug("Added {Count} temp files for cleanup", tempFiles.Count);
        }
        
        private void CleanupTempFiles()
        {
            foreach (var tempFile in _tempFilesToCleanup.Keys)
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
        
        /// <summary>
        /// Sanitizes a filename to remove invalid characters
        /// </summary>
        private string SanitizeFileName(string fileName)
        {
            // Get invalid characters for file names
            var invalidChars = Path.GetInvalidFileNameChars();
            var invalidCharsPattern = string.Join("", invalidChars.Select(c => Regex.Escape(c.ToString())));
            var pattern = $"[{invalidCharsPattern}]";
            
            // Replace invalid characters with underscore
            var sanitized = Regex.Replace(fileName, pattern, "_");
            
            // Also replace some additional problematic characters
            sanitized = sanitized.Replace(":", "_")
                               .Replace("<", "_")
                               .Replace(">", "_")
                               .Replace("\"", "_")
                               .Replace("/", "_")
                               .Replace("\\", "_")
                               .Replace("|", "_")
                               .Replace("?", "_")
                               .Replace("*", "_");
            
            // Trim dots and spaces from the ends
            sanitized = sanitized.Trim(' ', '.');
            
            // If the filename is empty or just dots/spaces, provide a default
            if (string.IsNullOrWhiteSpace(sanitized))
            {
                sanitized = "Document";
            }
            
            return sanitized;
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
                if (_tempFilesToCleanup.ContainsKey(fileItem.FilePath))
                {
                    try
                    {
                        if (File.Exists(fileItem.FilePath))
                        {
                            File.Delete(fileItem.FilePath);
                            _logger.Debug("Cleaned up removed temp file: {File}", fileItem.FilePath);
                        }
                        _tempFilesToCleanup.TryRemove(fileItem.FilePath, out _);
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
        private void ClearSelectedScope()
        {
            SelectedScope = null;
            // Don't clear the search text - user might want to search for another
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
                // since the database has changed - but only if user hasn't typed anything
                if (SaveQuotesMode && PendingFiles.Any() && string.IsNullOrWhiteSpace(CompanyNameInput))
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
            var window = new Views.EditScopesOfWorkWindow(_scopeOfWorkService)
            {
                Owner = Application.Current.MainWindow
            };
            
            if (window.ShowDialog() == true)
            {
                _logger.Information("Scopes of work were modified");
                
                // Reload scopes to reflect any changes
                LoadScopesOfWork();
                
                // If we're in Save Quotes mode, refresh the scope list
                if (SaveQuotesMode)
                {
                    // Clear any search to show all scopes
                    ScopeSearchText = "";
                    
                    // Reload recent scopes in case they were modified
                    LoadRecentScopes();
                }
            }
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
        
        [RelayCommand]
        private void ShowPerformanceMetrics()
        {
            try
            {
                var companyMetrics = _companyNameService.GetPerformanceSummary();
                var process = Process.GetCurrentProcess();
                var workingSet = process.WorkingSet64 / (1024 * 1024);
                var gcMemory = GC.GetTotalMemory(false) / (1024 * 1024);
                
                var message = $"DocHandler Performance Metrics\n\n" +
                             $"{companyMetrics}\n\n" +
                             $"Memory Usage:\n" +
                             $"  Working Set: {workingSet:N0} MB\n" +
                             $"  GC Memory: {gcMemory:N0} MB\n" +
                             $"  Thread Count: {process.Threads.Count}\n\n" +
                             $"Session Info:\n" +
                             $"  Files Processed: {_processedFileCount}\n" +
                             $"  Session Duration: {DateTime.Now - Process.GetCurrentProcess().StartTime:hh\\:mm\\:ss}";
                
                MessageBox.Show(message, "Performance Metrics", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to show performance metrics");
                MessageBox.Show("Failed to retrieve performance metrics.", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private async Task TestCompanyDetection()
        {
            // For debugging: manually test company detection
            if (!SaveQuotesMode)
            {
                MessageBox.Show("Please enable Save Quotes Mode first.", "Test Company Detection", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "Select document to test company detection",
                Filter = "All supported files|*.pdf;*.doc;*.docx;*.txt|PDF files|*.pdf|Word documents|*.doc;*.docx|Text files|*.txt|All files|*.*"
            };

            if (dialog.ShowDialog() == true)
            {
                // Clear any existing input
                CompanyNameInput = "";
                DetectedCompanyName = "";
                
                _logger.Information("Manual company detection test started for: {Path}", dialog.FileName);
                
                // Test the detection
                await ScanForCompanyName(dialog.FileName);
                
                // Show result
                var result = string.IsNullOrEmpty(DetectedCompanyName) 
                    ? "No company name detected." 
                    : $"Detected company: {DetectedCompanyName}";
                    
                MessageBox.Show($"Detection test completed.\n\n{result}", "Company Detection Test", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
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

            // Sanitize company name for use in filename
            companyName = SanitizeFileName(companyName);

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
                        // Scope is already sanitized in the database
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
                                companyName.Equals(SanitizeFileName(DetectedCompanyName), StringComparison.OrdinalIgnoreCase))
                            {
                                await _companyNameService.IncrementUsageCount(DetectedCompanyName);
                            }
                            
                            // Update scope usage
                            await _scopeOfWorkService.IncrementUsageCount(SelectedScope);
                            
                            // Add to company database if not already there (use original unsanitized name for database)
                            var originalCompanyName = !string.IsNullOrWhiteSpace(CompanyNameInput) 
                                ? CompanyNameInput.Trim() 
                                : DetectedCompanyName?.Trim();
                                
                            if (!string.IsNullOrWhiteSpace(originalCompanyName) &&
                                !_companyNameService.Companies.Any(c => 
                                    c.Name.Equals(originalCompanyName, StringComparison.OrdinalIgnoreCase)))
                            {
                                await _companyNameService.AddCompanyName(originalCompanyName);
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
                    
                    // Update processed file count for performance metrics
                    _processedFileCount += processedCount;
                    
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
            
            _logger.Debug("Processing single quote file: {File} ({Extension})", Path.GetFileName(inputPath), extension);
            
            // Add file existence check at the beginning
            if (!File.Exists(inputPath))
            {
                _logger.Warning("File no longer exists: {File}", inputPath);
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = "File no longer exists",
                    FailedFiles = { (inputPath, "File was deleted or moved") }
                };
            }
            
            // For PDFs, just copy
            if (extension == ".pdf")
            {
                try
                {
                    File.Copy(inputPath, outputPath, true);
                    _logger.Information("Copied PDF directly: {File}", Path.GetFileName(outputPath));
                    return new ProcessingResult 
                    { 
                        Success = true, 
                        SuccessfulFiles = { outputPath }
                    };
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to copy PDF");
                    return new ProcessingResult
                    {
                        Success = false,
                        ErrorMessage = $"Failed to copy PDF: {ex.Message}",
                        FailedFiles = { (inputPath, ex.Message) }
                    };
                }
            }
            
            // For Word documents, check if we have a cached PDF first
            if (extension == ".doc" || extension == ".docx")
            {
                // Lock to prevent race conditions if user rapidly processes files
                lock (_conversionLock)
                {
                    // Check for cached PDF from company detection
                    if (_companyNameService.TryGetCachedPdf(inputPath, out var cachedPdfPath) && cachedPdfPath != null)
                    {
                        try
                        {
                            _logger.Information("Using cached PDF from company detection for: {File}", Path.GetFileName(inputPath));
                            File.Copy(cachedPdfPath, outputPath, true);
                            
                            // Don't delete the cached PDF - let the cache manager handle it
                            return new ProcessingResult 
                            { 
                                Success = true, 
                                SuccessfulFiles = { outputPath }
                            };
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Failed to use cached PDF, will convert fresh");
                            // Fall through to fresh conversion
                        }
                    }
                }
                
                // No cache available, use session-aware service for conversion
                try
                {
                    _logger.Information("Converting Word document using session service: {File}", Path.GetFileName(inputPath));
                    var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                    
                    var conversionResult = await _sessionOfficeService.ConvertWordToPdf(inputPath, outputPath);
                    
                    stopwatch.Stop();
                    _logger.Information("Conversion completed in {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
                    
                    if (conversionResult.Success)
                    {
                        return new ProcessingResult 
                        { 
                            Success = true, 
                            SuccessfulFiles = { outputPath }
                        };
                    }
                    else
                    {
                        // If session service fails, try with regular service as fallback
                        _logger.Warning("Session service failed, trying fallback: {Error}", conversionResult.ErrorMessage);
                        
                        var fallbackResult = await _officeConversionService.ConvertWordToPdf(inputPath, outputPath);
                        
                        if (fallbackResult.Success)
                        {
                            return new ProcessingResult 
                            { 
                                Success = true, 
                                SuccessfulFiles = { outputPath }
                            };
                        }
                        else
                        {
                            return new ProcessingResult
                            {
                                Success = false,
                                ErrorMessage = fallbackResult.ErrorMessage ?? conversionResult.ErrorMessage,
                                FailedFiles = { (inputPath, fallbackResult.ErrorMessage ?? conversionResult.ErrorMessage ?? "Unknown error") }
                            };
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Word conversion failed");
                    return new ProcessingResult
                    {
                        Success = false,
                        ErrorMessage = $"Conversion failed: {ex.Message}",
                        FailedFiles = { (inputPath, ex.Message) }
                    };
                }
            }
            
            // For Excel, use existing service (can be optimized in future phase)
            if (extension == ".xls" || extension == ".xlsx")
            {
                try
                {
                    _logger.Information("Converting Excel document: {File}", Path.GetFileName(inputPath));
                    var conversionResult = await _officeConversionService.ConvertExcelToPdf(inputPath, outputPath);
                    
                    if (conversionResult.Success)
                    {
                        return new ProcessingResult 
                        { 
                            Success = true, 
                            SuccessfulFiles = { outputPath }
                        };
                    }
                    else
                    {
                        return new ProcessingResult
                        {
                            Success = false,
                            ErrorMessage = conversionResult.ErrorMessage,
                            FailedFiles = { (inputPath, conversionResult.ErrorMessage ?? "Unknown error") }
                        };
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Excel conversion failed");
                    return new ProcessingResult
                    {
                        Success = false,
                        ErrorMessage = $"Excel conversion failed: {ex.Message}",
                        FailedFiles = { (inputPath, ex.Message) }
                    };
                }
            }
            
            return new ProcessingResult
            {
                Success = false,
                ErrorMessage = $"Unsupported file type: {extension}"
            };
        }
        
        // Command handlers
        partial void OnCompanyNameInputChanged(string value)
        {
            UpdateUI();
        }
        
        partial void OnDetectedCompanyNameChanged(string value)
        {
            UpdateUI();
            OnPropertyChanged(nameof(CompanyNamePlaceholder));
        }
        
        public void Cleanup()
        {
            CleanupTempFiles();
            _scopeSearchTimer?.Stop();
            _scopeSearchTimer = null;
            _scopeSearchCancellation?.Cancel();
            _scopeSearchCancellation?.Dispose();
            _scopeSearchCancellation = null;
            
            // Cleanup session service
            try
            {
                _sessionOfficeService?.Dispose();
                _logger.Information("Session Office service disposed");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error disposing session Office service");
            }
            
            // Cleanup PDF cache
            try
            {
                _companyNameService?.CleanupPdfCache();
                _logger.Information("PDF cache cleaned up");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error cleaning PDF cache");
            }
            
            // Dispose other services
            _fileProcessingService?.Dispose();
            
            _logger.Information("MainViewModel cleanup completed");
        }
        
        public void SaveWindowState(double left, double top, double width, double height, string state)
        {
            // Validate window position before saving
            if (left >= SystemParameters.VirtualScreenLeft && 
                top >= SystemParameters.VirtualScreenTop &&
                left + width <= SystemParameters.VirtualScreenLeft + SystemParameters.VirtualScreenWidth &&
                top + height <= SystemParameters.VirtualScreenTop + SystemParameters.VirtualScreenHeight &&
                state != "Minimized") // Don't save minimized state
            {
                if (_configService.Config.RememberWindowPosition)
                {
                    _configService.UpdateWindowPosition(left, top, width, height, state);
                    _ = _configService.SaveConfiguration();
                }
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

        // Property for showing recent scopes
        private bool _showRecentScopes = false;
        public bool ShowRecentScopes
        {
            get => _showRecentScopes;
            set
            {
                if (SetProperty(ref _showRecentScopes, value))
                {
                    // Save the preference
                    _configService.Config.ShowRecentScopes = value;
                    _ = _configService.SaveConfiguration();
                }
            }
        }

        private bool _autoScanCompanyNames = true;
        public bool AutoScanCompanyNames
        {
            get => _autoScanCompanyNames;
            set
            {
                if (SetProperty(ref _autoScanCompanyNames, value))
                {
                    // Save the preference
                    _configService.Config.AutoScanCompanyNames = value;
                    _ = _configService.SaveConfiguration();
                    
                    // Clear detected company name if auto-scan is disabled
                    if (!value)
                    {
                        DetectedCompanyName = "";
                    }
                    
                    // Update placeholder text
                    OnPropertyChanged(nameof(CompanyNamePlaceholder));
                }
            }
        }

        public string CompanyNamePlaceholder
        {
            get
            {
                if (!AutoScanCompanyNames)
                {
                    return "Enter company name manually";
                }
                
                if (!string.IsNullOrEmpty(DetectedCompanyName))
                {
                    return DetectedCompanyName;
                }
                
                return "Scanning for company name...";
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