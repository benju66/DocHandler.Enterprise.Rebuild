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
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using DocHandler.Views;
using DocHandler.Models;
using DocHandler.ViewModels;
using Serilog;
using MessageBox = System.Windows.MessageBox;
using Application = System.Windows.Application;
using FolderBrowserDialog = Ookii.Dialogs.Wpf.VistaFolderBrowserDialog;
using System.Windows.Threading;

namespace DocHandler.ViewModels
{
    public partial class MainViewModel : ObservableObject, IDisposable
    {
        private readonly ILogger _logger;
        private readonly OptimizedFileProcessingService _fileProcessingService;
        private readonly ConfigurationService _configService;
        private readonly OfficeConversionService _officeConversionService;
        private readonly CompanyNameService _companyNameService;
        private readonly ScopeOfWorkService _scopeOfWorkService;
        private SessionAwareOfficeService? _sessionOfficeService;
        private SessionAwareExcelService? _sessionExcelService;
        private readonly PerformanceMonitor _performanceMonitor;
        private readonly PdfCacheService _pdfCacheService;
        private readonly ProcessManager _processManager;
        private readonly OfficeInstanceTracker _officeTracker;
        private SaveQuotesQueueService? _queueService;
        private readonly object _conversionLock = new object();
        private readonly ConcurrentDictionary<string, byte> _tempFilesToCleanup = new();
        
        // Queue service
        private QueueDetailsWindow _queueWindow;
        
        // Add concurrent scan protection
        private volatile int _activeScanCount = 0;
        private readonly SemaphoreSlim _scanSemaphore = new SemaphoreSlim(1, 1);
        private CancellationTokenSource? _currentScanCancellation;
        
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
                        
                        // Use centralized pre-warming strategy
                        PreWarmOfficeServicesForSaveQuotes();
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

        // Queue properties
        [ObservableProperty]
        private int _queueTotalCount;

        [ObservableProperty]
        private int _queueProcessedCount;

        [ObservableProperty]
        private bool _isQueueProcessing;

        [ObservableProperty]
        private string _queueStatusMessage = "Drop quote documents";

        // Add new properties for completion message
        [ObservableProperty]
        private string _queueCompletionMessage = "";

        [ObservableProperty]
        private bool _showQueueCompletionMessage = false;

        // Performance tracking
        private int _processedFileCount = 0;
        private PerformanceMetricsWindow _metricsWindow;

        // Animation state tracking
        private bool _isShowingSuccessAnimation;
        public bool IsShowingSuccessAnimation
        {
            get => _isShowingSuccessAnimation;
            set => SetProperty(ref _isShowingSuccessAnimation, value);
        }

        // Speed mode detection
        private DateTime _lastProcessTime = DateTime.MinValue;
        private bool _isSpeedMode = false;
        private const int SpeedModeThresholdSeconds = 3;

        // Recent locations
        public ObservableCollection<string> RecentLocations => 
            new ObservableCollection<string>(_configService.Config.RecentLocations);

        // .doc file scanning preferences
        private bool _scanCompanyNamesForDocFiles = false;
        public bool ScanCompanyNamesForDocFiles
        {
            get => _scanCompanyNamesForDocFiles;
            set
            {
                if (SetProperty(ref _scanCompanyNamesForDocFiles, value))
                {
                    _configService.Config.ScanCompanyNamesForDocFiles = value;
                    _ = _configService.SaveConfiguration();
                }
            }
        }

        private int _docFileSizeLimitMB = 10;
        public int DocFileSizeLimitMB
        {
            get => _docFileSizeLimitMB;
            set
            {
                if (SetProperty(ref _docFileSizeLimitMB, value))
                {
                    _configService.Config.DocFileSizeLimitMB = value;
                    _ = _configService.SaveConfiguration();
                    
                    // Update the company name service with the new limit
                    _companyNameService.UpdateDocFileSizeLimit(value);
                }
            }
        }
        
        public MainViewModel()
        {
            _logger = Log.ForContext<MainViewModel>();
            
            try 
            {
                // Initialize configuration first
                _configService = new ConfigurationService();
                
                // Initialize process manager 
                _processManager = new ProcessManager();
                
                // Initialize performance monitor
                _performanceMonitor = new PerformanceMonitor(_configService.Config.MemoryUsageLimitMB);
                
                // Initialize Office instance tracker FIRST (before any Office services)
                try
                {
                    _officeTracker = new OfficeInstanceTracker();
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to initialize OfficeInstanceTracker - continuing without it");
                    _officeTracker = null;
                }
                
                // Initialize PDF cache service
                _pdfCacheService = new PdfCacheService();
                
                // Initialize company detection service
                _companyNameService = new CompanyNameService();
                
                // Initialize scope service
                _scopeOfWorkService = new ScopeOfWorkService();
                

                
                // COORDINATED OFFICE SERVICES: Use only SessionAware services to reduce instance count
                // Remove individual OfficeConversionService creation - use SessionAware services for all operations
                try
                {
                    var timeoutTask = Task.Delay(10000); // 10 second timeout
                    var initTask = Task.Run(() => 
                    {
                        var tempOfficeService = new SessionAwareOfficeService();
                        var tempExcelService = new SessionAwareExcelService();
                        return (tempOfficeService, tempExcelService);
                    });
                    
                    if (Task.WhenAny(initTask, timeoutTask).Result == initTask && initTask.IsCompletedSuccessfully)
                    {
                        var result = initTask.Result;
                        _sessionOfficeService = result.tempOfficeService;
                        _sessionExcelService = result.tempExcelService;
                        _logger.Information("Shared SessionAware Office services initialized successfully");
                    }
                    else
                    {
                        // Initialize with null - will be initialized later when needed
                        _sessionOfficeService = null;
                        _sessionExcelService = null;
                        _logger.Warning("SessionAware Office services initialization timed out, will initialize when needed");
                    }
                }
                catch (Exception ex)
                {
                    _sessionOfficeService = null;
                    _sessionExcelService = null;
                    _logger.Warning(ex, "Failed to initialize SessionAware Office services, will initialize when needed");
                }
                
                // Initialize file processing service with shared session services
                try
                {
                    _fileProcessingService = new OptimizedFileProcessingService(
                        _configService, _pdfCacheService, _processManager, _officeTracker, 
                        _sessionOfficeService, _sessionExcelService);
                    _logger.Information("File processing service initialized with shared Office instances");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to initialize OptimizedFileProcessingService");
                    _fileProcessingService = null;
                }
                
                // Set Office services for company name detection if they were initialized
                if (_sessionOfficeService != null && _sessionExcelService != null)
                {
                    try
                    {
                        _companyNameService.SetOfficeServices(_sessionOfficeService, _sessionExcelService);
                        _logger.Information("Office services set for company name detection");
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to set Office services for company name detection");
                    }
                }
                
                // Queue service will be initialized lazily when needed
                
                // Subscribe to events
                SubscribeToEvents();
                
                // Initialize UI from configuration
                InitializeFromConfiguration();
                
                // Restore queue window if needed
                RestoreQueueWindowIfNeeded();
                
                // Load service data asynchronously after initial UI is shown
                _ = Task.Run(async () => await LoadServiceDataAsync());
                
                _logger.Information("MainViewModel initialized successfully");
            }
            catch (Exception ex)
            {
                _logger.Fatal(ex, "Failed to initialize MainViewModel");
                
                // Show error dialog but don't crash the app
                MessageBox.Show(
                    $"The application encountered an error during initialization:\n\n{ex.Message}\n\nSome features may not be available.",
                    "Initialization Warning",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                
                // Initialize with minimal functionality
                try
                {
                    if (_configService == null)
                        _configService = new ConfigurationService();
                    
                    InitializeFromConfiguration();
                }
                catch
                {
                    // Even minimal initialization failed
                    StatusMessage = "Initialization failed - limited functionality";
                }
            }
        }
        
        /// <summary>
        /// Loads service data asynchronously to avoid blocking UI thread during startup
        /// </summary>
        private async Task LoadServiceDataAsync()
        {
            try
            {
                _logger.Information("Starting async service data loading...");
                
                // Load data in parallel
                var loadTasks = new List<Task>();
                
                if (_companyNameService != null)
                {
                    loadTasks.Add(_companyNameService.LoadDataAsync());
                }
                
                if (_scopeOfWorkService != null)
                {
                    loadTasks.Add(_scopeOfWorkService.LoadDataAsync());
                }
                
                await Task.WhenAll(loadTasks);
                
                _logger.Information("Service data loaded successfully");
                
                // Update UI collections on dispatcher thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    try
                    {
                        // Reload scopes now that data is available
                        LoadScopesOfWork();
                        LoadRecentScopes();
                        
                        // If in Save Quotes Mode, ensure scopes are filtered
                        if (SaveQuotesMode)
                        {
                            FilterScopes();
                        }
                        
                        _logger.Information("UI collections updated with loaded data");
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to update UI collections after data load");
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load service data asynchronously");
            }
        }

        private void SubscribeToEvents()
        {
            try
            {
                // Queue events will be subscribed when queue service is lazily initialized
                
                // Performance monitor events - only subscribe if available
                if (_performanceMonitor != null)
                {
                    _performanceMonitor.MemoryPressureDetected += OnMemoryPressureDetected;
                }
                else
                {
                    _logger.Warning("Performance monitor is null - memory pressure detection will not be available");
                }
                
                // File collection changes
                PendingFiles.CollectionChanged += (s, e) => 
                {
                    Application.Current.Dispatcher.InvokeAsync(() => UpdateUI());
                    
                    if (SaveQuotesMode && e.NewItems != null && string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        foreach (FileItem item in e.NewItems)
                        {
                            StartCompanyNameScan(item.FilePath);
                            break;
                        }
                    }
                };
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error subscribing to events");
            }
        }

        private void InitializeFromConfiguration()
        {
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
            
            // Load .doc file scanning preferences
            ScanCompanyNamesForDocFiles = _configService.Config.ScanCompanyNamesForDocFiles;
            DocFileSizeLimitMB = _configService.Config.DocFileSizeLimitMB;
            
            // Initialize company name service with current size limit
            _companyNameService.UpdateDocFileSizeLimit(_configService.Config.DocFileSizeLimitMB);
            
                            // Check Office availability in background with timeout
                _ = Task.Run(async () => 
                {
                    try
                    {
                        var timeoutTask = Task.Delay(5000); // 5 second timeout
                        var checkTask = Task.Run(() => CheckOfficeAvailability());
                        
                        if (await Task.WhenAny(checkTask, timeoutTask) == timeoutTask)
                        {
                            _logger.Warning("Office availability check timed out");
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to check Office availability");
                    }
                });
            
            // Start periodic COM object monitoring
            StartComObjectMonitoring();
        }

        private void OnMemoryPressureDetected(object? sender, MemoryPressureEventArgs e)
        {
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                if (e.IsCritical)
                {
                    _logger.Warning("Critical memory pressure: {CurrentMB}MB", e.CurrentMemoryMB);
                    
                    MessageBox.Show(
                        $"Memory usage is critically high ({e.CurrentMemoryMB}MB). Consider closing other applications.",
                        "Memory Warning",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            });
        }
        
        private async Task ScanForCompanyName(string filePath, CancellationToken cancellationToken = default)
        {
            // Don't scan if auto-scan is disabled, user has already typed a company name, or other conditions
            if (!SaveQuotesMode || IsDetectingCompany || !string.IsNullOrWhiteSpace(CompanyNameInput) || !AutoScanCompanyNames) 
            {
                _logger.Debug("Skipping company name scan - SaveQuotesMode: {SaveQuotesMode}, IsDetecting: {IsDetecting}, HasInput: {HasInput}, AutoScan: {AutoScan}", 
                    SaveQuotesMode, IsDetectingCompany, !string.IsNullOrWhiteSpace(CompanyNameInput), AutoScanCompanyNames);
                return;
            }

            // Check if this is a .doc file and if .doc scanning is disabled
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            if (extension == ".doc" && !ScanCompanyNamesForDocFiles)
            {
                _logger.Information("Skipping company name scan for .doc file (disabled in settings): {Path}", filePath);
                
                // Show brief status message
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    StatusMessage = "Skipping company scan for .doc file";
                    
                    // Clear message after 2 seconds
                    var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        StatusMessage = "";
                        UpdateUI();
                    };
                    timer.Start();
                });
                
                return;
            }

            // Use semaphore to prevent concurrent scans and safer cancellation
            await _scanSemaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
            
            try
            {
                // Cancel any previous scan safely
                var previousCts = Interlocked.Exchange(ref _currentScanCancellation, null);
                if (previousCts != null)
                {
                    previousCts.Cancel();
                    // Dispose safely in background without blocking
                    _ = Task.Run(async () =>
                    {
                        await Task.Delay(200); // Give more time for cleanup
                        previousCts?.Dispose();
                    });
                }

                // Create a new cancellation token for this scan
                var localCts = new CancellationTokenSource();
                _currentScanCancellation = localCts;

                // Combine with method parameter and add timeout
                using var combinedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    cancellationToken, 
                    localCts.Token);

                combinedCts.CancelAfter(TimeSpan.FromSeconds(30));
            
            try
            {
                // Update UI on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    IsDetectingCompany = true;
                });
                
                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                _logger.Information("Starting optimized company name detection for: {Path}", filePath);
                
                // Create progress reporter for UI updates - use InvokeAsync consistently
                var progress = new Progress<int>(percentage =>
                {
                    // Use InvokeAsync instead of BeginInvoke for consistency
                    _ = Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        try
                        {
                            if (percentage <= 30)
                            {
                                StatusMessage = "Scanning document...";
                            }
                            else if (percentage <= 60)
                            {
                                StatusMessage = "Extracting text...";
                            }
                            else if (percentage <= 90)
                            {
                                StatusMessage = "Detecting company name...";
                            }
                            else
                            {
                                StatusMessage = "Finalizing detection...";
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error updating progress UI");
                        }
                    });
                });
                
                // Wrap the entire company name service operation in Task.Run with proper signature
                var detectedCompany = await Task.Run(async () =>
                {
                    // Now using the correct signature with progress reporting
                    return await _companyNameService.ScanDocumentForCompanyName(filePath, progress).ConfigureAwait(false);
                }, combinedCts.Token).ConfigureAwait(false);
                
                stopwatch.Stop();
                _logger.Information("Company detection completed in {ElapsedMs}ms", stopwatch.ElapsedMilliseconds);
                
                // Update UI on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
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
                });
            }
            catch (OperationCanceledException)
            {
                _logger.Warning("Company detection cancelled/timed out for: {Path}", filePath);
                // Show user-friendly message for timeout on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = "";
                    }
                    StatusMessage = "Company detection timed out";
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Company detection failed for: {Path}", filePath);
                // Clear any partial detection on error on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = "";
                    }
                    StatusMessage = "Company detection failed";
                });
            }
            finally
            {
                // Update UI on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    IsDetectingCompany = false;
                    StatusMessage = "";
                    _logger.Debug("Company name detection completed. DetectedCompanyName: '{DetectedName}'", 
                        DetectedCompanyName ?? "null");
                    UpdateUI(); // Ensure UI updates after detection completes
                });
            }
        }
        finally
        {
            _scanSemaphore.Release();
        }
    }
        
        // Add a new safe wrapper method
        private async Task ScanForCompanyNameSafely(string filePath)
        {
            // Prevent multiple concurrent scans
            if (_activeScanCount > 0)
            {
                _logger.Debug("Skipping scan - another scan is already in progress");
                return;
            }

            try
            {
                Interlocked.Increment(ref _activeScanCount);
                await ScanForCompanyName(filePath, CancellationToken.None);
            }
            finally
            {
                Interlocked.Decrement(ref _activeScanCount);
            }
        }

        private void StartCompanyNameScan(string filePath)
        {
            if (!SaveQuotesMode || !AutoScanCompanyNames || IsDetectingCompany || 
                !string.IsNullOrWhiteSpace(CompanyNameInput))
            {
                return;
            }
            
            // Cancel any existing scan
            _currentScanCancellation?.Cancel();
            _currentScanCancellation = new CancellationTokenSource();
            var cancellationToken = _currentScanCancellation.Token;
            
            // Run entirely in background
            Task.Run(async () =>
            {
                try
                {
                    // Update UI to show scanning
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        IsDetectingCompany = true;
                        DetectedCompanyName = "Scanning...";
                    });
                    
                    // Create progress reporter
                    var progress = new Progress<int>(percent =>
                    {
                        Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            if (percent < 100)
                            {
                                DetectedCompanyName = $"Scanning... {percent}%";
                            }
                        });
                    });
                    
                    // Perform detection with progress
                    var detectedCompany = await _companyNameService.ScanDocumentForCompanyName(
                        filePath, progress);
                    
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    
                    // Update UI with result
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (!string.IsNullOrWhiteSpace(detectedCompany))
                        {
                            DetectedCompanyName = detectedCompany;
                            CompanyNameInput = detectedCompany;
                            _logger.Information("Auto-detected company: {Company}", detectedCompany);
                        }
                        else
                        {
                            DetectedCompanyName = "No company detected";
                        }
                        
                        IsDetectingCompany = false;
                    });
                }
                catch (OperationCanceledException)
                {
                    _logger.Debug("Company name scan cancelled");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error during company name scan");
                    
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        DetectedCompanyName = "Detection failed";
                        IsDetectingCompany = false;
                    });
                }
            }, cancellationToken);
        }
        
        private void CheckOfficeAvailability()
        {
            try
            {
                if (_officeConversionService != null && !_officeConversionService.IsOfficeInstalled())
                {
                    _logger.Warning("Microsoft Office is not available - Word/Excel conversion features will be disabled");
                }
                else if (_officeConversionService == null)
                {
                    _logger.Warning("Office conversion service is not available - Office features will be limited");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error checking Office availability");
            }
        }
        
        /// <summary>
        /// Centralized pre-warming strategy for Office services when Save Quotes Mode is enabled
        /// </summary>
        private void PreWarmOfficeServicesForSaveQuotes()
        {
            if (_sessionOfficeService != null && _sessionExcelService != null)
            {
                // Check if we're on UI thread (which is STA)
                if (Application.Current.Dispatcher.CheckAccess())
                {
                    try
                    {
                        _logger.Information("Pre-warming Office services on UI thread (STA)...");
                        _logger.Information("Thread {ThreadId} apartment state: {ApartmentState}", 
                            System.Threading.Thread.CurrentThread.ManagedThreadId, 
                            System.Threading.Thread.CurrentThread.GetApartmentState());
                        
                        _sessionOfficeService.WarmUp();
                        _sessionExcelService.WarmUp();
                        
                        _logger.Information("✓ Office services pre-warmed successfully");
                        ComHelper.LogComObjectStats();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to pre-warm Office services");
                    }
                }
                else
                {
                    // Schedule on UI thread
                    _logger.Information("Scheduling Office warm-up on UI thread...");
                    Application.Current.Dispatcher.BeginInvoke(new Action(() =>
                    {
                        PreWarmOfficeServicesForSaveQuotes();
                    }));
                }
            }
            else
            {
                _logger.Debug("Office services not available for pre-warming - they will be initialized when first needed");
            }
        }
        
        private SaveQuotesQueueService GetOrCreateQueueService()
        {
            if (_queueService == null)
            {
                try
                {
                    _logger.Information("Initializing SaveQuotesQueueService on first use");
                    _queueService = new SaveQuotesQueueService(_configService, _pdfCacheService, _processManager, _fileProcessingService);
                    
                    // Subscribe to queue events
                    _queueService.ProgressChanged += OnQueueProgressChanged;
                    _queueService.ItemCompleted += OnQueueItemCompleted;
                    _queueService.QueueEmpty += OnQueueEmpty;
                    _queueService.StatusMessageChanged += OnQueueStatusMessageChanged;
                    
                    _logger.Information("SaveQuotesQueueService initialized successfully");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to initialize SaveQuotesQueueService");
                    throw new InvalidOperationException("Queue service initialization failed. Save Quotes Mode requires the queue service to function.", ex);
                }
            }
            return _queueService;
        }
        
        private void StartComObjectMonitoring()
        {
            // Log COM object statistics every 5 minutes during application runtime
            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(5)
            };
            
            timer.Tick += (s, e) =>
            {
                Task.Run(() =>
                {
                    try
                    {
                        ComHelper.LogComObjectStats();
                        _processManager?.LogProcessInfo();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error during COM object monitoring");
                    }
                });
            };
            
            timer.Start();
            _logger.Information("COM object monitoring started - statistics will be logged every 5 minutes");
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
        private async void FilterScopes()
        {
            // CRITICAL FIX: Capture all UI values on the UI thread FIRST
            var uiValues = await Application.Current.Dispatcher.InvokeAsync(() => new
            {
                SearchTerm = ScopeSearchText?.Trim() ?? "",
                CurrentSelection = SelectedScope,
                ScopesOfWorkList = ScopesOfWork.ToList() // Create a safe copy
            });
            
            // Build filtered list without clearing to minimize UI disruption
            var filteredList = new List<string>();
            
            if (string.IsNullOrWhiteSpace(uiValues.SearchTerm))
            {
                // No search term - show all scopes
                filteredList.AddRange(uiValues.ScopesOfWorkList);
            }
            else
            {
                // Fuzzy search implementation
                var searchWords = uiValues.SearchTerm.ToLowerInvariant().Split(new[] { ' ', '-' }, StringSplitOptions.RemoveEmptyEntries);
                
                var scoredScopes = new List<(string scope, double score)>();
                
                foreach (var scope in uiValues.ScopesOfWorkList)
                {
                    double score = CalculateFuzzyScore(scope, uiValues.SearchTerm, searchWords);
                    if (score > 0)
                    {
                        scoredScopes.Add((scope, score));
                    }
                }
                
                // Sort by score (highest first) and add to filtered list
                filteredList.AddRange(scoredScopes.OrderByDescending(x => x.score).Select(x => x.scope));
            }
            
            // Update the filtered collection efficiently using ASYNC dispatcher
            await Application.Current.Dispatcher.InvokeAsync(() =>
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
                if (uiValues.CurrentSelection != null)
                {
                    if (FilteredScopesOfWork.Contains(uiValues.CurrentSelection))
                    {
                        // Keep existing selection if it's still in filtered results
                        if (SelectedScope != uiValues.CurrentSelection)
                        {
                            SelectedScope = uiValues.CurrentSelection;
                        }
                    }
                    else if (string.IsNullOrWhiteSpace(uiValues.SearchTerm))
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
            // Ensure this method is always called on the UI thread
            if (!Application.Current.Dispatcher.CheckAccess())
            {
                Application.Current.Dispatcher.InvokeAsync(() => UpdateUI());
                return;
            }
            
            if (SaveQuotesMode)
            {
                // Need both a scope and either typed or detected company name
                var hasCompanyName = !string.IsNullOrWhiteSpace(CompanyNameInput) || 
                                   !string.IsNullOrWhiteSpace(DetectedCompanyName);
                
                // Only enable processing if all files are valid
                var allFilesValid = PendingFiles.All(f => f.ValidationStatus == ValidationStatus.Valid);
                
                CanProcess = PendingFiles.Count > 0 && 
                            allFilesValid &&
                            !IsProcessing && 
                            !string.IsNullOrEmpty(SelectedScope) && 
                            hasCompanyName;
                
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
                // Only enable processing if all files are valid
                var allFilesValid = PendingFiles.All(f => f.ValidationStatus == ValidationStatus.Valid);
                
                CanProcess = PendingFiles.Count > 0 && allFilesValid && !IsProcessing;
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
            var addedFiles = new List<FileItem>();
            
            foreach (var file in filePaths)
            {
                // Quick validation only - existence and not already added
                if (!File.Exists(file))
                {
                    _logger.Warning("File does not exist: {FilePath}", file);
                    continue;
                }
                
                if (PendingFiles.Any(f => f.FilePath == file))
                {
                    _logger.Information("File already in list: {FilePath}", file);
                    continue;
                }
                
                try
                {
                    var fileInfo = new FileInfo(file);
                    var fileItem = new FileItem
                    {
                        FilePath = file,
                        FileName = Path.GetFileName(file),
                        FileSize = fileInfo.Length,
                        FileType = Path.GetExtension(file).ToUpperInvariant().TrimStart('.'),
                        ValidationStatus = ValidationStatus.Pending
                    };
                    
                    PendingFiles.Add(fileItem);
                    addedFiles.Add(fileItem);
                    
                    _logger.Debug("File added instantly: {FileName}", fileItem.FileName);
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to add file: {FilePath}", file);
                }
            }
            
            // Validate files in background
            if (addedFiles.Any())
            {
                _ = Task.Run(async () => await ValidateFilesAsync(addedFiles));
            }
            
            UpdateUI();
        }

        private async Task ValidateFilesAsync(List<FileItem> files)
        {
            foreach (var fileItem in files)
            {
                try
                {
                    // Update status to validating
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        fileItem.ValidationStatus = ValidationStatus.Validating;
                    });
                    
                    // Perform thorough validation
                    var validationResult = await Task.Run(() => 
                        DocHandler.Helpers.FileValidator.ValidateFile(fileItem.FilePath));
                    
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (validationResult.IsValid)
                        {
                            fileItem.ValidationStatus = ValidationStatus.Valid;
                            
                            // If in Save Quotes mode and no company input yet, start scan
                            if (SaveQuotesMode && string.IsNullOrWhiteSpace(CompanyNameInput) && 
                                files.IndexOf(fileItem) == 0) // Only scan first file
                            {
                                StartCompanyNameScan(fileItem.FilePath);
                            }
                        }
                        else
                        {
                            fileItem.ValidationStatus = ValidationStatus.Invalid;
                            fileItem.ValidationError = validationResult.ErrorMessage;
                            
                            _logger.Warning("File validation failed: {File} - {Error}", 
                                fileItem.FileName, validationResult.ErrorMessage);
                        }
                        
                        UpdateUI();
                    });
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error validating file: {File}", fileItem.FileName);
                    
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        fileItem.ValidationStatus = ValidationStatus.Invalid;
                        fileItem.ValidationError = "Validation error occurred";
                    });
                }
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
        
        private async Task CleanupTempFiles()
        {
            var tempFilesToRemove = _tempFilesToCleanup.Keys.ToList();
            
            if (tempFilesToRemove.Count == 0)
                return;
            
            // Move file operations to background thread to avoid blocking UI
            await Task.Run(() =>
            {
                foreach (var tempFile in tempFilesToRemove)
                {
                    try
                    {
                        if (File.Exists(tempFile))
                        {
                            File.Delete(tempFile);
                            _logger.Debug("Cleaned up temp file: {File}", tempFile);
                        }
                        _tempFilesToCleanup.TryRemove(tempFile, out _);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to cleanup temp file: {File}", tempFile);
                    }
                }
            }).ConfigureAwait(false);
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
            // Quick validation on UI thread
            var pendingValidation = PendingFiles.Where(f => 
                f.ValidationStatus == ValidationStatus.Pending || 
                f.ValidationStatus == ValidationStatus.Validating).ToList();
                
            if (pendingValidation.Any())
            {
                MessageBox.Show("Please wait for file validation to complete.", 
                    "Validation in Progress", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Information);
                return;
            }
            
            // Remove any invalid files on UI thread
            var invalidFiles = PendingFiles.Where(f => 
                f.ValidationStatus == ValidationStatus.Invalid).ToList();
                
            if (invalidFiles.Any())
            {
                var message = $"The following files are invalid and will be skipped:\n\n" +
                    string.Join("\n", invalidFiles.Select(f => $"• {f.FileName}: {f.ValidationError}"));
                    
                MessageBox.Show(message, "Invalid Files", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                    
                foreach (var invalid in invalidFiles)
                {
                    PendingFiles.Remove(invalid);
                }
                
                if (!PendingFiles.Any()) return;
            }
            
            // Handle Save Quotes mode
            if (SaveQuotesMode)
            {
                await ProcessSaveQuotes();
                return;
            }

            // Capture UI state immediately on UI thread
            var hasFiles = PendingFiles.Any();
            var fileCount = PendingFiles.Count;
            var filePaths = PendingFiles.Select(f => f.FilePath).ToList();
            var openFolderAfterProcessing = this.OpenFolderAfterProcessing;

            if (!hasFiles)
            {
                StatusMessage = "No files selected";
                return;
            }

            // Set processing state on UI thread
            IsProcessing = true;
            StatusMessage = fileCount > 1 ? "Merging and processing files..." : "Processing file...";

            // Move heavy processing to background thread
            await Task.Run(async () =>
            {
                await ProcessFilesBackground(filePaths, fileCount, openFolderAfterProcessing);
            });
        }

        private async Task ProcessFilesBackground(List<string> filePaths, int fileCount, bool openFolderAfterProcessing)
        {
            try
            {
                // Check if file processing service is available
                if (_fileProcessingService == null)
                {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        StatusMessage = "File processing service is not available";
                        MessageBox.Show(
                            "The file processing service could not be initialized. Please check the logs for more information.",
                            "Service Unavailable",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    });
                    return;
                }
                
                var outputDir = _configService.Config.DefaultSaveLocation;

                // Create output folder with timestamp
                outputDir = _fileProcessingService.CreateOutputFolder(outputDir);

                var result = await _fileProcessingService.ProcessFiles(filePaths, outputDir, ConvertOfficeToPdf).ConfigureAwait(false);

                if (result.Success)
                {
                    // Update UI on UI thread
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (result.IsMerged)
                        {
                            StatusMessage = $"Successfully merged {filePaths.Count} files into {Path.GetFileName(result.SuccessfulFiles.First())}";
                        }
                        else
                        {
                            StatusMessage = $"Successfully processed {result.SuccessfulFiles.Count} file(s)";
                        }

                        // Clear the file list after successful processing
                        PendingFiles.Clear();
                    });

                    if (result.IsMerged)
                    {
                        _logger.Information("Files merged successfully");
                    }
                    else
                    {
                        _logger.Information("Files processed successfully");
                    }
                    
                    // Clean up any temp files
                    await CleanupTempFiles().ConfigureAwait(false);

                    // Update configuration with recent location
                    _configService.AddRecentLocation(outputDir);

                    // Open the output folder if preference is set
                    if (openFolderAfterProcessing)
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
                    
                    // Update UI on UI thread
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        StatusMessage = $"Error: {errorMessage}";
                    });
                    
                    _logger.Error("File processing failed: {Error}", errorMessage);

                    if (result.FailedFiles.Any())
                    {
                        var failedFilesList = string.Join("\n", result.FailedFiles.Select(f => 
                            $"• {Path.GetFileName(f.FilePath)}: {f.Error}"));
                        
                        await Application.Current.Dispatcher.InvokeAsync(() =>
                        {
                            MessageBox.Show(
                                $"The following files could not be processed:\n\n{failedFilesList}",
                                "Processing Errors",
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                // Update UI on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    StatusMessage = $"Error: {ex.Message}";
                });
                
                _logger.Error(ex, "Unexpected error during file processing");
                
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show(
                        $"An unexpected error occurred:\n\n{ex.Message}",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                });
            }
            finally
            {
                // Ensure all UI updates happen on the UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    IsProcessing = false;
                    ProgressValue = 0;
                    UpdateUI();
                });
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
                
                // Clean up any cached PDF from company detection
                _companyNameService.RemoveCachedPdf(fileItem.FilePath);
                
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
            await _scopeOfWorkService.ClearRecentScopes().ConfigureAwait(false);
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
                    
                    // Use the safe wrapper instead of fire-and-forget
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await ScanForCompanyNameSafely(firstFile.FilePath);
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Error during rescan after company names edit");
                        }
                    });
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
        private async Task RunQueueDiagnosticAsync()
        {
            try
            {
                StatusMessage = "Running queue diagnostic...";
                
                var diagnosticResult = await Task.Run(async () => 
                {
                    return await QuickDiagnostic.RunQueueDiagnosticAsync();
                });
                
                // Show results in a message box
                MessageBox.Show(diagnosticResult, "Queue Processing Diagnostic Results", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                StatusMessage = "Diagnostic completed";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during queue diagnostic");
                MessageBox.Show($"Diagnostic failed: {ex.Message}", "Diagnostic Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusMessage = "Diagnostic failed";
            }
        }

        [RelayCommand]
        private void OpenSettings()
        {
            var settingsViewModel = new SettingsViewModel(
                _configService, 
                _companyNameService, 
                _scopeOfWorkService);
                
            var settingsWindow = new Views.SettingsWindow(settingsViewModel)
            {
                Owner = Application.Current.MainWindow
            };
            
            if (settingsWindow.ShowDialog() == true)
            {
                // Reload settings that affect the UI
                IsDarkMode = _configService.Config.Theme == "Dark";
                SaveQuotesMode = _configService.Config.SaveQuotesMode;
                ShowRecentScopes = _configService.Config.ShowRecentScopes;
                OpenFolderAfterProcessing = _configService.Config.OpenFolderAfterProcessing ?? true;
                AutoScanCompanyNames = _configService.Config.AutoScanCompanyNames;
                ScanCompanyNamesForDocFiles = _configService.Config.ScanCompanyNamesForDocFiles;
                DocFileSizeLimitMB = _configService.Config.DocFileSizeLimitMB;
                
                // Update queue service with new parallel limit
                _queueService?.UpdateMaxConcurrency(_configService.Config.MaxParallelProcessing);
                
                // Apply theme change
                if (_configService.Config.Theme == "Dark")
                {
                    ModernWpf.ThemeManager.Current.ApplicationTheme = ModernWpf.ApplicationTheme.Dark;
                }
                else
                {
                    ModernWpf.ThemeManager.Current.ApplicationTheme = ModernWpf.ApplicationTheme.Light;
                }
                
                _logger.Information("Settings updated from preferences window");
            }
        }
        
        [RelayCommand]
        private async Task ShowPerformanceMetricsAsync()
        {
            try
            {
                _logger.Information("ShowPerformanceMetricsAsync called");
                
                // Close existing window if open
                if (_metricsWindow != null && _metricsWindow.IsLoaded)
                {
                    _logger.Information("Activating existing metrics window");
                    _metricsWindow.Activate();
                    return;
                }
                
                _logger.Information("Creating new PerformanceMetricsWindow");
                
                // Create and show new non-blocking window
                _metricsWindow = new PerformanceMetricsWindow(this);
                _metricsWindow.Owner = Application.Current.MainWindow;
                _metricsWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                _metricsWindow.Closed += (s, e) => _metricsWindow = null;
                _metricsWindow.Show(); // Non-blocking
                
                _logger.Information("Performance metrics window opened successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to show performance metrics window");
                MessageBox.Show($"Failed to open performance metrics window:\n\n{ex.Message}\n\nInner Exception: {ex.InnerException?.Message}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Collects comprehensive performance metrics asynchronously
        /// </summary>
        public async Task<string> CollectPerformanceMetricsAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    var documentMetrics = _performanceMonitor.GetDocumentProcessingPerformanceSummary();
                    var companyMetrics = _companyNameService.GetPerformanceSummary();
                    var memoryInfo = _performanceMonitor.GetMemoryInfo();
                    var systemInfo = _performanceMonitor.GetSystemPerformanceInfo();
                    
                    var process = Process.GetCurrentProcess();
                    var workingSet = process.WorkingSet64 / (1024 * 1024);
                    var gcMemory = GC.GetTotalMemory(false) / (1024 * 1024);
                    
                    var metrics = new StringBuilder();
                    metrics.AppendLine("DocHandler Performance Metrics");
                    metrics.AppendLine(new string('=', 50));
                    metrics.AppendLine();
                    metrics.AppendLine($"=== {documentMetrics}");
                    metrics.AppendLine($"=== Company Detection Performance ===");
                    metrics.AppendLine(companyMetrics);
                    metrics.AppendLine();
                    metrics.AppendLine("=== Memory Usage ===");
                    metrics.AppendLine($"  Working Set: {workingSet:N0} MB");
                    metrics.AppendLine($"  GC Memory: {gcMemory:N0} MB");
                    metrics.AppendLine($"  Peak Memory: {memoryInfo.PeakMemoryMB} MB");
                    metrics.AppendLine($"  Memory Growth: {memoryInfo.MemoryGrowthMB} MB");
                    metrics.AppendLine($"  Thread Count: {process.Threads.Count}");
                    metrics.AppendLine();
                    metrics.AppendLine("=== System Performance ===");
                    metrics.AppendLine($"  CPU Usage: {systemInfo.CpuUsagePercent:F1}%");
                    metrics.AppendLine($"  Available RAM: {systemInfo.AvailableMemoryMB} MB");
                    metrics.AppendLine();
                    metrics.AppendLine("=== Session Info ===");
                    metrics.AppendLine($"  Files Processed: {_processedFileCount}");
                    metrics.AppendLine($"  Session Duration: {DateTime.Now - process.StartTime:hh\\:mm\\:ss}");
                    metrics.AppendLine($"  Current Time: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    
                    return metrics.ToString();
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to collect performance metrics");
                    return "Error collecting metrics: " + ex.Message;
                }
            });
        }

        [RelayCommand]
        private async Task TestCompanyDetection()
        {
            if (PendingFiles.Count == 0)
            {
                MessageBox.Show("Please add at least one file to test company detection.", "No Files", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            IsDetectingCompany = true;
            StatusMessage = "Testing company detection...";
            
            try
            {
                foreach (var file in PendingFiles)
                {
                    // Use the updated signature with proper cancellation
                    await ScanForCompanyName(file.FilePath, CancellationToken.None);
                }
                
                if (string.IsNullOrEmpty(DetectedCompanyName))
                {
                    MessageBox.Show(
                        "No company name was detected in the selected files.\n\n" +
                        "This could mean:\n" +
                        "• The files don't contain recognizable company names\n" +
                        "• The company names aren't in the configured list\n" +
                        "• The files are image-based PDFs or unsupported formats\n\n" +
                        "You can manually enter the company name or edit the company list.",
                        "No Company Detected",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show(
                        $"Company detected: {DetectedCompanyName}",
                        "Company Detection Test",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during company detection test");
                MessageBox.Show($"An error occurred during company detection:\n\n{ex.Message}", 
                    "Test Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsDetectingCompany = false;
                UpdateUI();
            }
        }

        [RelayCommand]
        private async Task ShowComStatsAsync()
        {
            try
            {
                // Run on background thread like other diagnostics
                await Task.Run(() =>
                {
                    // Get current stats
                    var stats = ComHelper.GetComObjectSummary();
                    
                    // Log to file for record keeping
                    _logger.Information("COM Statistics requested by user");
                    ComHelper.LogComObjectStats();
                    
                    // Build the message
                    var message = new StringBuilder();
                    message.AppendLine("COM Object Lifecycle Statistics");
                    message.AppendLine("================================\n");
                    
                    message.AppendLine($"Total Created:  {stats.TotalCreated,8}");
                    message.AppendLine($"Total Released: {stats.TotalReleased,8}");
                    message.AppendLine($"Net Objects:    {stats.NetObjects,8}");
                    message.AppendLine();
                    
                    if (stats.NetObjects == 0)
                    {
                        message.AppendLine("✓ All COM objects properly released!");
                        message.AppendLine("  No memory leaks detected.");
                    }
                    else
                    {
                        message.AppendLine("⚠ Potential COM object leaks detected!");
                        message.AppendLine("\nUnreleased objects by type:");
                        
                        foreach (var kvp in stats.ObjectStats.Where(s => s.Value.Net > 0).OrderByDescending(s => s.Value.Net))
                        {
                            var stat = kvp.Value;
                            message.AppendLine($"\n  {stat.ObjectType} ({stat.Context}):");
                            message.AppendLine($"    Created:  {stat.Created}");
                            message.AppendLine($"    Released: {stat.Released}");
                            message.AppendLine($"    Leaked:   {stat.Net}");
                        }
                        
                        message.AppendLine("\nRecommendation: Check the log file for details.");
                    }
                    
                    message.AppendLine($"\nStats tracking started: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                    
                    // Show the stats window on UI thread
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show(
                            Application.Current.MainWindow,
                            message.ToString(),
                            "COM Object Statistics",
                            MessageBoxButton.OK,
                            stats.NetObjects > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
                    });
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error showing COM statistics");
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show(
                        $"Error displaying COM statistics: {ex.Message}",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                });
            }
        }

        [RelayCommand]
        private async Task ResetComStatsAsync()
        {
            var result = MessageBox.Show(
                "This will reset all COM object tracking statistics.\n\n" +
                "This is useful when you want to start fresh monitoring from a specific point.\n\n" +
                "Continue?",
                "Reset COM Statistics",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
                
            if (result == MessageBoxResult.Yes)
            {
                await Task.Run(() =>
                {
                    ComHelper.ResetStats();
                    _logger.Information("COM statistics reset by user");
                });
                
                MessageBox.Show(
                    "COM object statistics have been reset.\n\n" +
                    "All counters are now at zero.",
                    "Statistics Reset",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
        }

        // Add method to restore queue window on startup
        private void RestoreQueueWindowIfNeeded()
        {
            if (_configService.Config.QueueWindowIsOpen && _configService.Config.RestoreQueueWindowOnStartup)
            {
                // Delay restoration to ensure main window is fully loaded
                var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (_queueWindow == null && Application.Current.MainWindow?.IsLoaded == true)
                        {
                            ShowQueueWindow();
                            _logger.Information("Restored queue window from previous session");
                        }
                    });
                };
                timer.Start();
            }
        }

        // Add helper method to ensure window is visible
        private void EnsureWindowIsOnScreen(Window window)
        {
            var workingArea = SystemParameters.WorkArea;
            
            if (window.Left < 0) window.Left = 0;
            if (window.Top < 0) window.Top = 0;
            
            if (window.Left + window.Width > workingArea.Width)
                window.Left = workingArea.Width - window.Width;
            
            if (window.Top + window.Height > workingArea.Height)
                window.Top = workingArea.Height - window.Height;
        }

        [RelayCommand]
        private void ShowQueueWindow()
        {
            if (_queueWindow == null || !_queueWindow.IsLoaded)
            {
                // Ensure queue service is initialized before showing window
                var queueService = GetOrCreateQueueService();
                
                _queueWindow = new QueueDetailsWindow(queueService, _configService)
                {
                    Owner = Application.Current.MainWindow
                };
                
                // Restore window position
                if (_configService.Config.QueueWindowLeft.HasValue && _configService.Config.QueueWindowTop.HasValue)
                {
                    _queueWindow.Left = _configService.Config.QueueWindowLeft.Value;
                    _queueWindow.Top = _configService.Config.QueueWindowTop.Value;
                    
                    // Ensure window is on screen
                    EnsureWindowIsOnScreen(_queueWindow);
                }
                
                // Subscribe to closed event
                _queueWindow.Closed += (s, e) =>
                {
                    _queueWindow = null;
                    _configService.Config.QueueWindowIsOpen = false;
                    _ = _configService.SaveConfiguration();
                };
                
                _queueWindow.Show();
                
                // Mark as open
                _configService.Config.QueueWindowIsOpen = true;
                _ = _configService.SaveConfiguration();
            }
            else
            {
                _queueWindow.Activate();
            }
        }

        // Queue event handlers
        private void OnQueueProgressChanged(object? sender, SaveQuoteProgressEventArgs e)
        {
            Application.Current.Dispatcher.InvokeAsync(() =>
            {
                QueueTotalCount = e.TotalCount;
                QueueProcessedCount = e.ProcessedCount;
                IsQueueProcessing = e.IsProcessing;
            });
        }

        private void OnQueueItemCompleted(object? sender, SaveQuoteCompletedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (e.Success)
                {
                    QueueStatusMessage = "Successfully saved!";
                    
                    // Reset message after 1 second
                    var timer = new DispatcherTimer 
                    { 
                        Interval = TimeSpan.FromSeconds(1) 
                    };
                    timer.Tick += (s, args) =>
                    {
                        timer.Stop();
                        UpdateQueueStatusMessage();
                    };
                    timer.Start();
                }
            });
        }

        private void OnQueueEmpty(object? sender, EventArgs e)
        {
            Application.Current.Dispatcher.Invoke(async () =>
            {
                if (_queueService == null) return; // Safety check
                
                var totalCount = _queueService.TotalCount;
                var failedCount = _queueService.FailedCount;
                var successCount = totalCount - failedCount;
                
                if (failedCount > 0)
                {
                    MessageBox.Show(
                        $"{successCount} quotes saved, {failedCount} failed - see queue for details",
                        "Processing Complete",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
                
                // Show success animation if no failures and not in speed mode
                if (failedCount == 0 && !_isSpeedMode)
                {
                    if (Application.Current.MainWindow is MainWindow mainWindow)
                    {
                        await mainWindow.ShowSaveQuotesSuccessAnimation(totalCount);
                    }
                }
                
                // Set completion message for middle status bar
                if (failedCount > 0)
                {
                    QueueCompletionMessage = $"Saved {successCount} of {totalCount}";
                }
                else
                {
                    QueueCompletionMessage = $"Saved {totalCount} of {totalCount}";
                }
                
                ShowQueueCompletionMessage = true;
                
                // Hide completion message after 4 seconds
                var completionTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(4) };
                completionTimer.Tick += (s, args) =>
                {
                    completionTimer.Stop();
                    ShowQueueCompletionMessage = false;
                };
                completionTimer.Start();
                
                QueueStatusMessage = "All quotes saved";
                
                // Reset to idle after 2 seconds
                var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                timer.Tick += (s, args) =>
                {
                    timer.Stop();
                    QueueStatusMessage = "Drop quote documents";
                };
                timer.Start();
            });
        }

        private void OnQueueStatusMessageChanged(object? sender, string message)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                QueueStatusMessage = message;
            });
        }

        private void UpdateQueueStatusMessage()
        {
            if (IsQueueProcessing)
            {
                QueueStatusMessage = "Processing queue...";
            }
            else if (QueueTotalCount > 0)
            {
                QueueStatusMessage = $"{QueueTotalCount} item(s) in queue";
            }
            else
            {
                QueueStatusMessage = "Drop quote documents";
            }
        }

        private async Task ProcessSaveQuotes()
        {
            // Quick validation and UI state capture on UI thread
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

            // Capture UI state immediately on UI thread
            var selectedScope = this.SelectedScope;
            var pendingFilesList = this.PendingFiles.ToList(); // Create a copy
            var sessionSaveLocation = this.SessionSaveLocation;
            var fileCount = pendingFilesList.Count;

            // Move heavy processing to background thread
            await Task.Run(async () =>
            {
                await ProcessSaveQuotesBackground(selectedScope, companyName, pendingFilesList, sessionSaveLocation, fileCount);
            });
        }

        private async Task ProcessSaveQuotesBackground(string selectedScope, string companyName, List<FileItem> pendingFilesList, string sessionSaveLocation, int fileCount)
        {
            try
            {
                // Get or create queue service
                SaveQuotesQueueService queueService;
                try
                {
                    queueService = GetOrCreateQueueService();
                }
                catch (Exception ex)
                {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        StatusMessage = "Queue service initialization failed";
                        MessageBox.Show(
                            $"The queue service could not be initialized:\n\n{ex.Message}",
                            "Service Unavailable",
                            MessageBoxButton.OK,
                            MessageBoxImage.Error);
                    });
                    return;
                }
                
                // Sanitize company name for use in filename
                companyName = SanitizeFileName(companyName);

                var outputDir = !string.IsNullOrEmpty(sessionSaveLocation) 
                    ? sessionSaveLocation 
                    : _configService.Config.DefaultSaveLocation;
                
                // Add to queue instead of processing directly
                foreach (var file in pendingFilesList)
                {
                    queueService.AddToQueue(file, selectedScope, companyName, outputDir);
                }
                
                // Clear UI immediately on UI thread
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    PendingFiles.Clear();
                    CompanyNameInput = "";
                    DetectedCompanyName = "";
                    SelectedScope = null;
                    StatusMessage = $"Added {fileCount} quote{(fileCount > 1 ? "s" : "")} to queue";
                });
                
                // Start processing if not already running
                if (!queueService.IsProcessing)
                {
                    _ = Task.Run(async () =>
                    {
                        try
                        {
                            await queueService.StartProcessingAsync();
                        }
                        catch (Exception ex)
                        {
                            _logger.Error(ex, "Failed to start queue processing");
                            
                            await Application.Current.Dispatcher.InvokeAsync(() =>
                            {
                                StatusMessage = "Queue processing failed";
                                MessageBox.Show(
                                    $"Failed to start queue processing:\n\n{ex.Message}",
                                    "Queue Error",
                                    MessageBoxButton.OK,
                                    MessageBoxImage.Error);
                            });
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Unexpected error during save quotes processing");
                
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    StatusMessage = $"Error: {ex.Message}";
                    MessageBox.Show(
                        $"An unexpected error occurred:\n\n{ex.Message}",
                        "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                });
            }
        }

        private void OpenOutputFolder(string outputDir)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = outputDir,
                    UseShellExecute = true,
                    Verb = "open"
                });
                _logger.Information("Opened output folder: {Dir}", outputDir);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to open output folder: {Dir}", outputDir);
            }
        }
        
        private async Task<ProcessingResult> ProcessSingleQuoteFile(string inputPath, string outputPath)
        {
            var extension = Path.GetExtension(inputPath).ToLowerInvariant();
            
            _logger.Debug("Processing single quote file: {File} ({Extension})", Path.GetFileName(inputPath), extension);
            
            // Check if we have a cached PDF from company detection
            var cachedPdf = _companyNameService.GetCachedPdfPath(inputPath);
            if (!string.IsNullOrEmpty(cachedPdf) && File.Exists(cachedPdf))
            {
                try
                {
                    _logger.Information("Using cached PDF from company detection for {File}", Path.GetFileName(inputPath));
                    
                    // Just copy the already-converted PDF
                    File.Copy(cachedPdf, outputPath, overwrite: true);
                    
                    // Clean up the cached PDF
                    _companyNameService.RemoveCachedPdf(inputPath);
                    
                    return new ProcessingResult 
                    { 
                        Success = true,
                        SuccessfulFiles = { outputPath }
                    };
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to use cached PDF, will reconvert");
                    // Fall through to normal conversion
                }
            }
            
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
                // Check if Office conversion service is available
                if (_officeConversionService == null)
                {
                    _logger.Error("Office conversion service is not available");
                    return new ProcessingResult
                    {
                        Success = false,
                        ErrorMessage = "Microsoft Office is required to convert Word documents to PDF",
                        FailedFiles = { (inputPath, "Office not available") }
                    };
                }
                
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
                    
                    var conversionResult = await _sessionOfficeService.ConvertWordToPdf(inputPath, outputPath).ConfigureAwait(false);
                    
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
                        
                        var fallbackResult = await _officeConversionService.ConvertWordToPdf(inputPath, outputPath).ConfigureAwait(false);
                        
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
                    var conversionResult = await _officeConversionService.ConvertExcelToPdf(inputPath, outputPath).ConfigureAwait(false);
                    
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
        
        public async void Cleanup()
        {
            try
            {
                _logger.Information("MainViewModel cleanup started");
                
                // Close queue window if open
                _queueWindow?.Close();
                
                // Cancel any active scans safely
                var currentScan = Interlocked.Exchange(ref _currentScanCancellation, null);
                if (currentScan != null)
                {
                    currentScan.Cancel();
                    // Give some time for operations to complete
                    await Task.Delay(500);
                    currentScan.Dispose();
                }
                
                // Dispose semaphore
                _scanSemaphore?.Dispose();
                
                await CleanupTempFiles().ConfigureAwait(false);
                
                // Stop and dispose scope search timer
                _scopeSearchTimer?.Stop();
                _scopeSearchTimer = null;
                
                // Handle scope search cancellation safely
                var scopeSearchCancellation = Interlocked.Exchange(ref _scopeSearchCancellation, null);
                if (scopeSearchCancellation != null)
                {
                    scopeSearchCancellation.Cancel();
                    await Task.Delay(100);
                    scopeSearchCancellation.Dispose();
                }
            
                // Dispose queue service first (it uses file processing service)
                try
                {
                    _queueService?.Dispose();
                    _logger.Information("Queue service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error disposing queue service");
                }
                
                // Dispose file processing service
                try
                {
                    _fileProcessingService?.Dispose();
                    _logger.Information("File processing service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error disposing file processing service");
                }
                
                // Dispose company name service
                try
                {
                    _companyNameService?.Dispose();
                    _logger.Information("Company name service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error disposing company name service");
                }
                
                // Dispose Office services last (after services that use them)
                if (_sessionOfficeService != null)
                {
                    try
                    {
                        _sessionOfficeService.Dispose();
                        _logger.Information("Session Office service disposed");
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error disposing Session Office service");
                    }
                }
                
                if (_sessionExcelService != null)
                {
                    try
                    {
                        _sessionExcelService.Dispose();
                        _logger.Information("Session Excel service disposed");
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error disposing Session Excel service");
                    }
                }
                
                if (_officeConversionService != null)
                {
                    try
                    {
                        _officeConversionService.Dispose();
                        _logger.Information("Office conversion service disposed");
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Error disposing Office conversion service");
                    }
                }
                
                // Force COM cleanup before disposing other services
                ComHelper.ForceComCleanup("MainViewModelCleanup");
                
                // Log final COM stats
                ComHelper.LogComObjectStats();
                
                // Dispose performance monitor
                _performanceMonitor?.Dispose();
                _logger.Information("Performance monitor disposed");
                
                // Dispose office instance tracker (before process manager)
                _officeTracker?.Dispose();
                _logger.Information("Office instance tracker disposed");
                
                // Dispose process manager last
                _processManager?.Dispose();
                _logger.Information("Process manager disposed");
                
                _logger.Information("MainViewModel cleanup completed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during MainViewModel cleanup");
            }
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

        // Property for binding
        public bool RestoreQueueWindowOnStartup
        {
            get => _configService.Config.RestoreQueueWindowOnStartup;
            set
            {
                if (_configService.Config.RestoreQueueWindowOnStartup != value)
                {
                    _configService.Config.RestoreQueueWindowOnStartup = value;
                    _ = _configService.SaveConfiguration();
                    OnPropertyChanged();
                }
            }
        }

        // Property for queue window state indicator
        public bool QueueWindowIsOpen => _configService.Config.QueueWindowIsOpen;

        // Command to reset position
        [RelayCommand]
        private void ResetQueueWindowPosition()
        {
            _configService.Config.QueueWindowLeft = null;
            _configService.Config.QueueWindowTop = null;
            _configService.Config.QueueWindowWidth = 600;
            _configService.Config.QueueWindowHeight = 400;
            _ = _configService.SaveConfiguration();
            
            // If window is open, move it to center
            if (_queueWindow?.IsLoaded == true)
            {
                _queueWindow.Width = 600;
                _queueWindow.Height = 400;
                _queueWindow.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            }
            
            MessageBox.Show("Queue window position has been reset.", "Reset Complete", 
                            MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #region IDisposable Implementation
        
        private bool _disposed = false;
        
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Unsubscribe from all events to prevent memory leaks
                    UnsubscribeFromEvents();
                    
                    // Call the existing cleanup method
                    Cleanup();
                }
                
                _disposed = true;
            }
        }
        
        private void UnsubscribeFromEvents()
        {
            try
            {
                // Unsubscribe from queue service events
                if (_queueService != null)
                {
                    _queueService.ProgressChanged -= OnQueueProgressChanged;
                    _queueService.ItemCompleted -= OnQueueItemCompleted;
                    _queueService.QueueEmpty -= OnQueueEmpty;
                    _queueService.StatusMessageChanged -= OnQueueStatusMessageChanged;
                }
                
                // Unsubscribe from performance monitor events
                if (_performanceMonitor != null)
                {
                    _performanceMonitor.MemoryPressureDetected -= OnMemoryPressureDetected;
                }
                
                // Unsubscribe from file collection changes
                if (PendingFiles != null)
                {
                    PendingFiles.CollectionChanged -= null; // This was an anonymous method, so we can't unsubscribe cleanly
                }
                
                _logger.Information("Event handlers unsubscribed");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error unsubscribing from events");
            }
        }
        
        ~MainViewModel()
        {
            Dispose(false);
        }
        
        #endregion
    }
}
