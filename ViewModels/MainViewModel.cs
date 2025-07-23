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
using DocHandler.Services.Configuration;
using DocHandler.Views;
using DocHandler.Models;
using DocHandler.ViewModels;
using Serilog;
using MessageBox = System.Windows.MessageBox;
using Application = System.Windows.Application;
using FolderBrowserDialog = Ookii.Dialogs.Wpf.VistaFolderBrowserDialog;
using System.Windows.Threading;
using Microsoft.Extensions.DependencyInjection;

namespace DocHandler.ViewModels
{
    public partial class MainViewModel : ObservableObject, IDisposable
    {
        private readonly ILogger _logger;
        private readonly IOptimizedFileProcessingService _fileProcessingService;
        private readonly ConfigurationService _configService;
        private readonly IOfficeConversionService _officeConversionService;
        private readonly ICompanyNameService _companyNameService;
        private readonly IScopeOfWorkService _scopeOfWorkService;
        
        // Business Logic Services (Phase 2 Milestone 2)
        private readonly IFileValidationService _fileValidationService;
        private readonly IEnhancedCompanyDetectionService _companyDetectionService;
        private readonly IScopeManagementService _scopeManagementService;
        private readonly IDocumentWorkflowService _documentWorkflowService;
        private readonly IUIStateService _uiStateService;

        // Advanced Mode UI Framework (Phase 2 Milestone 2 - Day 5)
        private readonly IAdvancedModeUIProvider _advancedModeUIProvider;
        private readonly IDynamicMenuBuilder _dynamicMenuBuilder;
        private readonly IAdvancedModeUIManager _advancedModeUIManager;

        // Public access for MainWindow integration
        public IAdvancedModeUIManager AdvancedModeUIManager => _advancedModeUIManager;
        private SessionAwareOfficeService? _sessionOfficeService;
        private SessionAwareExcelService? _sessionExcelService;
        private readonly PerformanceMonitor _performanceMonitor;
        private readonly PdfCacheService _pdfCacheService;
        private readonly ProcessManager _processManager;
        private SaveQuotesQueueService? _queueService;
        private readonly OfficeHealthMonitor _healthMonitor;
        private ErrorRecoveryService? _errorRecoveryService;
        private TelemetryService? _telemetryService;
        private PdfOperationsService? _pdfOperationsService;
        private IModeManager? _modeManager;
        private readonly object _conversionLock = new object();
        private readonly ConcurrentDictionary<string, byte> _tempFilesToCleanup = new();
        
        // Queue service
        private QueueDetailsWindow _queueWindow;
        
        // Add concurrent scan protection
        private volatile int _activeScanCount = 0;
        private readonly SemaphoreSlim _scanSemaphore = new SemaphoreSlim(1, 1);
        private CancellationTokenSource? _currentScanCancellation;
        
        // Add for proper lifecycle management
        private readonly CancellationTokenSource _viewModelCts = new();
        private CancellationTokenSource? _filterCts;
        
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
                        _ = _uiStateService.UpdateStatusAsync("Save Quotes Mode: Drop quote documents");
                        SessionSaveLocation = _configService.Config.DefaultSaveLocation;
                        
                        // Pre-warming removed - instances created on-demand for better memory management
                    }
                    else
                    {
                        _ = _uiStateService.UpdateStatusAsync("Drop files here to begin");
                        SelectedScope = null;
                        CompanyNameInput = "";
                        DetectedCompanyName = "";
                    }
                }
            }
        }

        // Advanced Mode UI Framework Properties (Phase 2 Milestone 2 - Day 5)
        public string AdvancedCurrentMode => _advancedModeUIManager?.CurrentMode ?? "default";
        
        [ObservableProperty]
        private string _advancedCurrentModeDisplayName = "Standard Processing";
        
        [ObservableProperty]
        private bool _isAdvancedModeUIEnabled = true;

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
        
        // Cache for scope parts to improve search performance
        private readonly Dictionary<string, (string code, string description)> _scopePartsCache = new();

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
        
        public MainViewModel(
            IOptimizedFileProcessingService fileProcessingService,
            IConfigurationService configService,
            IOfficeConversionService officeConversionService,
            ICompanyNameService companyNameService,
            IScopeOfWorkService scopeOfWorkService,
            ErrorRecoveryService errorRecoveryService,
            TelemetryService telemetryService,
            PdfOperationsService pdfOperationsService,
            PerformanceMonitor performanceMonitor,
            PdfCacheService pdfCacheService,
            IProcessManager processManager,
            OfficeHealthMonitor healthMonitor,
            IFileValidationService fileValidationService,
            IEnhancedCompanyDetectionService companyDetectionService,
            IScopeManagementService scopeManagementService,
            IDocumentWorkflowService documentWorkflowService,
            IUIStateService uiStateService,
            IAdvancedModeUIProvider advancedModeUIProvider,
            IDynamicMenuBuilder dynamicMenuBuilder,
            IAdvancedModeUIManager advancedModeUIManager,
            IModeManager? modeManager = null)
        {
            try
            {
                _logger = Log.ForContext<MainViewModel>();
                _logger.Information("MainViewModel initialization started with dependency injection");

                // Initialize services from DI
                _fileProcessingService = fileProcessingService ?? throw new ArgumentNullException(nameof(fileProcessingService));
                _configService = (ConfigurationService)configService ?? throw new ArgumentNullException(nameof(configService));
                _companyNameService = companyNameService ?? throw new ArgumentNullException(nameof(companyNameService));
                _scopeOfWorkService = scopeOfWorkService ?? throw new ArgumentNullException(nameof(scopeOfWorkService));
                _errorRecoveryService = errorRecoveryService ?? throw new ArgumentNullException(nameof(errorRecoveryService));
                _telemetryService = telemetryService ?? throw new ArgumentNullException(nameof(telemetryService));
                _pdfOperationsService = pdfOperationsService ?? throw new ArgumentNullException(nameof(pdfOperationsService));
                _performanceMonitor = performanceMonitor ?? throw new ArgumentNullException(nameof(performanceMonitor));
                _pdfCacheService = pdfCacheService ?? throw new ArgumentNullException(nameof(pdfCacheService));
                _processManager = (ProcessManager)processManager ?? throw new ArgumentNullException(nameof(processManager));
                _healthMonitor = healthMonitor ?? throw new ArgumentNullException(nameof(healthMonitor));
                _modeManager = modeManager;
                
                // Initialize business logic services (Phase 2 Milestone 2)
                _fileValidationService = fileValidationService ?? throw new ArgumentNullException(nameof(fileValidationService));
                _companyDetectionService = companyDetectionService ?? throw new ArgumentNullException(nameof(companyDetectionService));
                _scopeManagementService = scopeManagementService ?? throw new ArgumentNullException(nameof(scopeManagementService));
                _documentWorkflowService = documentWorkflowService ?? throw new ArgumentNullException(nameof(documentWorkflowService));
                _uiStateService = uiStateService ?? throw new ArgumentNullException(nameof(uiStateService));

                // Initialize advanced mode UI framework (Phase 2 Milestone 2 - Day 5)
                _advancedModeUIProvider = advancedModeUIProvider ?? throw new ArgumentNullException(nameof(advancedModeUIProvider));
                _dynamicMenuBuilder = dynamicMenuBuilder ?? throw new ArgumentNullException(nameof(dynamicMenuBuilder));
                _advancedModeUIManager = advancedModeUIManager ?? throw new ArgumentNullException(nameof(advancedModeUIManager));
                
                        _officeConversionService = officeConversionService;
                
                // CRITICAL MEMORY FIX: Don't use shared session services - create on demand only
                _sessionOfficeService = null; // Always null - will create ReliableOfficeConverter on demand
                _sessionExcelService = null; // Always null - will create ReliableOfficeConverter on demand

                _ = _uiStateService.UpdateStatusAsync("Ready");
                
                // Initialize from configuration
                InitializeFromConfiguration();
                
                // Load service data asynchronously after UI is ready
                _ = Task.Run(async () => 
                {
                    await LoadServiceDataAsync();
                    await Application.Current.Dispatcher.InvokeAsync(() => 
                    {
                        // Reload UI collections after data is loaded
                        LoadScopesOfWork();
                        LoadRecentScopes();
                        if (SaveQuotesMode)
                        {
                            FilterScopes();
                        }
                    });
                });
                
                _logger.Information("MainViewModel initialized successfully with dependency injection");
            }
            catch (ConfigurationException configEx)
            {
                _logger.Error(configEx, "Configuration error during MainViewModel initialization");
                
                var errorInfo = _errorRecoveryService.CreateErrorInfo(configEx, "MainViewModel initialization");
                _ = _uiStateService.ShowWarningAsync(
                    errorInfo.Title,
                    $"{errorInfo.Message}\n\n{errorInfo.RecoveryGuidance}");
                
                // Try to initialize with default configuration
                try
                {
                    _configService = new ConfigurationService();
                    InitializeFromConfiguration();
                    _ = _uiStateService.UpdateStatusAsync("Initialized with default settings");
                }
                catch (Exception fallbackEx)
                {
                    _logger.Fatal(fallbackEx, "Failed to initialize even with defaults");
                    _ = _uiStateService.UpdateStatusAsync("Initialization failed - limited functionality");
                }
            }
            catch (ServiceException serviceEx)
            {
                _logger.Error(serviceEx, "Service initialization error");
                
                var errorInfo = _errorRecoveryService.CreateErrorInfo(serviceEx, "Service initialization");
                _ = _uiStateService.ShowWarningAsync(
                    errorInfo.Title,
                    $"{errorInfo.Message}\n\n{errorInfo.RecoveryGuidance}");
                
                // Initialize minimal services
                InitializeMinimalServices();
            }
            catch (Exception ex)
            {
                _logger.Fatal(ex, "Unexpected error during MainViewModel initialization");
                
                // Use error recovery service for unknown exceptions
                var errorInfo = _errorRecoveryService.CreateErrorInfo(ex, "MainViewModel initialization");
                _ = _uiStateService.ShowErrorAsync(
                    errorInfo.Title,
                    $"{errorInfo.Message}\n\n{errorInfo.RecoveryGuidance}");
                
                // Try minimal initialization
                InitializeMinimalServices();
            }
        }

        /// <summary>
        /// Initialize minimal services when full initialization fails
        /// </summary>
        private void InitializeMinimalServices()
        {
            try
            {
                // Can only initialize non-readonly services here
                if (_errorRecoveryService == null) 
                    _errorRecoveryService = new ErrorRecoveryService();
                    
                StatusMessage = "Running with limited functionality";
                
                _logger.Information("Minimal services initialized successfully");
            }
            catch (Exception ex)
            {
                _logger.Fatal(ex, "Even minimal service initialization failed");
                StatusMessage = "Critical initialization failure";
            }
        }

        /// <summary>
        /// Try to initialize mode manager from DI container (graceful fallback)
        /// </summary>
        private void TryInitializeModeManager()
        {
            try
            {
                _logger.Information("Mode system initialization deferred - using legacy SaveQuotes implementation");
                // Mode manager integration will be completed in future milestones
                // For now, maintain 100% backward compatibility with existing SaveQuotes functionality
                _modeManager = null;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to initialize mode manager - falling back to legacy implementation");
                _modeManager = null;
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
                
                await Task.WhenAll(loadTasks).ConfigureAwait(false);
                
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

                // Initialize Advanced Mode UI Framework (Phase 2 Milestone 2 - Day 5)
                await InitializeAdvancedModeUIAsync().ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load service data asynchronously");
            }
        }

        /// <summary>
        /// Initialize Advanced Mode UI Framework (Phase 2 Milestone 2 - Day 5)
        /// </summary>
        private async Task InitializeAdvancedModeUIAsync()
        {
            try
            {
                _logger.Information("Initializing Advanced Mode UI Framework...");

                // Subscribe to mode change events
                _advancedModeUIManager.ModeChanged += OnAdvancedModeChanged;

                // Initialize with the current mode based on SaveQuotesMode
                var initialMode = SaveQuotesMode ? "SaveQuotes" : "default";
                await _advancedModeUIManager.InitializeAsync(initialMode);

                _logger.Information("Advanced Mode UI Framework initialized successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to initialize Advanced Mode UI Framework");
            }
        }

        /// <summary>
        /// Handle advanced mode changes
        /// </summary>
        private void OnAdvancedModeChanged(object? sender, AdvancedModeChangedEventArgs e)
        {
            try
            {
                Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    // Update the display name
                    AdvancedCurrentModeDisplayName = e.UICustomization?.DisplayName ?? e.CurrentMode;
                    
                    // Sync with existing SaveQuotesMode if needed
                    var shouldBeSaveQuotesMode = e.CurrentMode == "SaveQuotes";
                    if (SaveQuotesMode != shouldBeSaveQuotesMode)
                    {
                        _saveQuotesMode = shouldBeSaveQuotesMode; // Set backing field directly to avoid recursion
                        OnPropertyChanged(nameof(SaveQuotesMode));
                        
                        // Update UI accordingly
                        UpdateUI();
                    }

                    // Notify property changed for UI updates
                    OnPropertyChanged(nameof(AdvancedCurrentMode));
                    
                    _logger.Information("Mode changed from {PreviousMode} to {CurrentMode}", 
                        e.PreviousMode, e.CurrentMode);
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error handling advanced mode change");
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
            Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                if (e.IsCritical)
                {
                    _logger.Warning("Critical memory pressure: {CurrentMB}MB", e.CurrentMemoryMB);
                    
                    await _uiStateService.ShowWarningAsync(
                        "Memory Warning",
                        $"Memory usage is critically high ({e.CurrentMemoryMB}MB). Consider closing other applications.");
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
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    await _uiStateService.UpdateStatusAsync("Skipping company scan for .doc file");
                    
                    // Clear message after 2 seconds
                    var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                    timer.Tick += async (s, args) =>
                    {
                        timer.Stop();
                        await _uiStateService.UpdateStatusAsync("");
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
                    _ = Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        try
                        {
                            if (percentage <= 30)
                            {
                                await _uiStateService.UpdateStatusAsync("Scanning document...");
                            }
                            else if (percentage <= 60)
                            {
                                await _uiStateService.UpdateStatusAsync("Extracting text...");
                            }
                            else if (percentage <= 90)
                            {
                                await _uiStateService.UpdateStatusAsync("Detecting company name...");
                            }
                            else
                            {
                                await _uiStateService.UpdateStatusAsync("Finalizing detection...");
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
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = "";
                    }
                    await _uiStateService.UpdateStatusAsync("Company detection timed out");
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Company detection failed for: {Path}", filePath);
                // Clear any partial detection on error on UI thread
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (string.IsNullOrWhiteSpace(CompanyNameInput))
                    {
                        DetectedCompanyName = "";
                    }
                    await _uiStateService.UpdateStatusAsync("Company detection failed");
                });
            }
            finally
            {
                // Update UI on UI thread
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    IsDetectingCompany = false;
                    await _uiStateService.UpdateStatusAsync("");
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
            
            // Run entirely in background using the new company detection service (Phase 2 Milestone 2)
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
                    
                    // Use the new company detection service
                    var request = new CompanyDetectionRequest { FilePath = filePath };
                    var companyResult = await _companyDetectionService.DetectCompanyAsync(request.FilePath);
                    var detectedCompany = companyResult?.DetectedCompany;
                    
                    if (cancellationToken.IsCancellationRequested)
                        return;
                    
                    // Update UI with result
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        if (!string.IsNullOrWhiteSpace(detectedCompany))
                        {
                            DetectedCompanyName = detectedCompany;
                            CompanyNameInput = detectedCompany;
                            _logger.Information("Auto-detected company using new service: {Company}", detectedCompany);
                        }
                        else
                        {
                            DetectedCompanyName = "No company detected";
                        }
                        
                        IsDetectingCompany = false;
                        UpdateUI();
                    });
                }
                catch (OperationCanceledException)
                {
                    _logger.Debug("Company name scan cancelled");
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        IsDetectingCompany = false;
                        DetectedCompanyName = "";
                    });
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
        
        private SaveQuotesQueueService GetOrCreateQueueService()
        {
            if (_queueService == null)
            {
                try
                {
                    _logger.Information("Initializing SaveQuotesQueueService on first use");
                    _queueService = new SaveQuotesQueueService(_configService, _pdfCacheService, _processManager, (OptimizedFileProcessingService)_fileProcessingService);
                    
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
            try
            {
                // Clear cache when reloading scopes
                _scopePartsCache.Clear();
                
                var scopes = _scopeOfWorkService.Scopes
                    .Select(s => _scopeOfWorkService.GetFormattedScope(s))
                    .ToList();
                
                ScopesOfWork.Clear();
                foreach (var scope in scopes)
                {
                                            ScopesOfWork.Add(scope);
                }
                
                // Initialize filtered list with all scopes
                FilterScopes();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load scopes of work");
            }
        }

        private async Task LoadRecentScopesAsync(CancellationToken cancellationToken = default)
        {
            try
            {
                // Combine our disposal token with the provided one
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(_viewModelCts.Token, cancellationToken);
                var token = cts.Token;
                
                // Use the new scope management service
                var recentScopes = await _scopeManagementService.GetRecentScopesAsync().ConfigureAwait(false);
                
                // Check for disposal/cancellation before UI updates
                if (_disposed || token.IsCancellationRequested) return;
                
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    RecentScopes.Clear();
                    foreach (var scope in recentScopes)
                    {
                        if (token.IsCancellationRequested) break;
                        RecentScopes.Add(scope.Name);
                    }
                }, DispatcherPriority.DataBind, token);
                
                _logger.Debug("Loaded {Count} recent scopes using new service", recentScopes.Count);
            }
            catch (OperationCanceledException) when (_viewModelCts.Token.IsCancellationRequested)
            {
                _logger.Debug("LoadRecentScopes cancelled due to disposal");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load recent scopes using new service");
                
                if (_disposed || cancellationToken.IsCancellationRequested) return;
                
                try
                {
                    await Application.Current.Dispatcher.InvokeAsync(() =>
                    {
                        RecentScopes.Clear();
                        var fallbackScopes = _scopeOfWorkService.RecentScopes.Take(10).ToList();
                        foreach (var scope in fallbackScopes)
                        {
                            RecentScopes.Add(scope);
                        }
                    }, DispatcherPriority.DataBind);
                }
                catch (Exception fallbackEx)
                {
                    _logger.Error(fallbackEx, "Fallback scope loading also failed");
                }
            }
        }

        // Keep the sync wrapper for backward compatibility
        private void LoadRecentScopes()
        {
            _ = LoadRecentScopesAsync().ContinueWith(task =>
            {
                if (task.IsFaulted)
                {
                    _logger.Error(task.Exception, "Background scope loading failed");
                }
            }, TaskScheduler.Default);
        }

        // Enhanced fuzzy search implementation using ScopeManagementService (Phase 2 Milestone 2)
        private async void FilterScopes()
        {
            try
            {
                // Cancel any previous search
                _filterCts?.Cancel();
                _filterCts = new CancellationTokenSource();
                var token = _filterCts.Token;
                
                // Capture search term immediately
                var searchTerm = ScopeSearchText?.Trim() ?? "";
                
                // Debounce: wait 300ms for user to stop typing
                try
                {
                    await Task.Delay(300, token);
                }
                catch (OperationCanceledException)
                {
                    return; // User typed again, cancel this search
                }
                
                // Use the new scope management service for searching
                var filteredScopes = await _scopeManagementService.FilterScopesAsync(searchTerm).ConfigureAwait(false);
                
                if (token.IsCancellationRequested) return;
                
                // Update UI with filtered results
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    var currentSelection = SelectedScope;
                    
                    // Update collection efficiently
                    FilteredScopesOfWork.Clear();
                    foreach (var scope in filteredScopes)
                    {
                        FilteredScopesOfWork.Add(scope.Name);
                    }
                    
                    // Preserve selection if it's in the filtered results
                    if (currentSelection != null && FilteredScopesOfWork.Contains(currentSelection))
                    {
                        SelectedScope = currentSelection;
                    }
                    else if (string.IsNullOrWhiteSpace(searchTerm))
                    {
                        SelectedScope = null;
                    }
                });
                
                _logger.Debug("Scope filtering completed using new service: {ResultCount} results for '{SearchTerm}'", 
                    filteredScopes.Count, searchTerm);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Scope filtering failed using new service");
                
                // Fallback to showing all scopes on error
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    FilteredScopesOfWork.Clear();
                    foreach (var scope in ScopesOfWork)
                    {
                        FilteredScopesOfWork.Add(scope);
                    }
                });
            }
        }
        
        private async Task UpdateFilteredScopes(List<string> filteredList, string searchTerm)
        {
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var currentSelection = SelectedScope;
                
                // Update collection efficiently
                FilteredScopesOfWork.Clear();
                foreach (var scope in filteredList)
                {
                    FilteredScopesOfWork.Add(scope);
                }
                
                // Preserve selection if it's in the filtered results
                if (currentSelection != null && FilteredScopesOfWork.Contains(currentSelection))
                {
                    SelectedScope = currentSelection;
                }
                else if (string.IsNullOrWhiteSpace(searchTerm))
                {
                    SelectedScope = null;
                }
            });
        }
        
        private async Task PerformFuzzySearch(string searchTerm)
        {
            // CRITICAL FIX: Capture all UI values on the UI thread FIRST
            var uiValues = await Application.Current.Dispatcher.InvokeAsync(() => new
            {
                SearchTerm = searchTerm,
                CurrentSelection = SelectedScope,
                ScopesOfWorkList = ScopesOfWork.ToList() // Create a safe copy
            });
            
            // Build filtered list without clearing to minimize UI disruption
            var filteredList = new List<string>();
            
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
            var searchTermLower = searchTerm.ToLowerInvariant();
            
            // Use cached scope parts if available, otherwise compute and cache
            if (!_scopePartsCache.TryGetValue(scope, out var parts))
            {
                var dashIndex = scope.IndexOf(" - ", StringComparison.Ordinal);
                parts = (
                    code: dashIndex > 0 ? scope.Substring(0, dashIndex) : "",
                    description: dashIndex > 0 ? scope.Substring(dashIndex + 3) : scope
                );
                _scopePartsCache[scope] = parts;
            }
            
            double score = 0;
            
            // 1. Exact match (highest score)
            if (scope.Equals(searchTerm, StringComparison.OrdinalIgnoreCase))
            {
                return 100;
            }
            
            // 2. Exact code match
            if (!string.IsNullOrEmpty(parts.code) && parts.code.Equals(searchTerm, StringComparison.OrdinalIgnoreCase))
            {
                return 90;
            }
            
            // 3. Code starts with search term
            if (!string.IsNullOrEmpty(parts.code) && parts.code.StartsWith(searchTerm, StringComparison.OrdinalIgnoreCase))
            {
                score += 80 - (parts.code.Length - searchTermLower.Length); // Closer matches score higher
            }
            
            // 4. Code contains search term
            else if (!string.IsNullOrEmpty(parts.code) && parts.code.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 60;
            }
            
            // 5. Description exact match
            if (parts.description.Equals(searchTerm, StringComparison.OrdinalIgnoreCase))
            {
                score += 85;
            }
            
            // 6. Description starts with search term
            else if (parts.description.StartsWith(searchTerm, StringComparison.OrdinalIgnoreCase))
            {
                score += 70;
            }
            
            // 7. Full scope contains exact search term
            else if (scope.IndexOf(searchTerm, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                score += 50;
                // Bonus if it's at a word boundary
                if (Regex.IsMatch(scope, $@"\b{Regex.Escape(searchTerm)}\b", RegexOptions.IgnoreCase))
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
                    if (scope.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0)
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
                var scopeWords = scope.Split(new[] { ' ', '-', ',', '.' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var scopeWord in scopeWords)
                {
                    if (scopeWord.Equals(searchWord, StringComparison.OrdinalIgnoreCase))
                    {
                        score += 35; // Exact word match
                    }
                    else if (scopeWord.StartsWith(searchWord, StringComparison.OrdinalIgnoreCase))
                    {
                        score += 25; // Word starts with search
                    }
                    else if (scopeWord.IndexOf(searchWord, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        score += 15; // Word contains search
                    }
                }
            }
            
            // 10. Fuzzy matching for typos (Levenshtein distance) - only for medium length searches
            if (score == 0 && searchTermLower.Length >= 3 && searchTermLower.Length <= 10) // Limit to reasonable lengths
            {
                // Check description words for close matches
                var descWords = parts.description.Split(new[] { ' ', '-', ',', '.' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var word in descWords)
                {
                    // Skip if word lengths are too different (optimization)
                    if (Math.Abs(word.Length - searchTermLower.Length) > 3) continue;
                    
                    var distance = LevenshteinDistance(searchTermLower, word.ToLowerInvariant());
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
            // Delegate complex UI state logic to UIStateService
            var context = new UIStateContext
            {
                SaveQuotesMode = SaveQuotesMode,
                PendingFileCount = PendingFiles.Count,
                AllFilesValid = PendingFiles.All(f => f.ValidationStatus == ValidationStatus.Valid),
                IsProcessing = IsProcessing,
                SelectedScope = SelectedScope,
                CompanyNameInput = CompanyNameInput,
                DetectedCompanyName = DetectedCompanyName
            };

            // Use UIStateService for complex state management
            _ = Task.Run(async () => await _uiStateService.RefreshUIStateAsync(context));
        }
        
        public void AddFiles(string[] filePaths)
        {
            // Use the new FileValidationService for dropped file validation (Phase 2 Milestone 2)
            _ = Task.Run(async () => await AddFilesAsync(filePaths));
        }

        /// <summary>
        /// Asynchronously adds and validates files using the new validation service (Phase 2 Milestone 2)
        /// </summary>
        private async Task AddFilesAsync(string[] filePaths)
        {
            try
            {
                _logger.Information("Adding {FileCount} dropped files", filePaths.Length);

                // Filter out files that are already added
                var newFilePaths = filePaths.Where(path => 
                    !PendingFiles.Any(f => f.FilePath == path)).ToArray();

                if (!newFilePaths.Any())
                {
                    _logger.Information("All dropped files are already in the list");
                    return;
                }

                // Use the new file validation service to validate dropped files
                var validatedFiles = await _fileValidationService.ValidateDroppedFilesAsync(newFilePaths);

                // Update UI with the validated files
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var fileItem in validatedFiles)
                    {
                        PendingFiles.Add(fileItem);
                        _logger.Debug("File added and validated: {FileName} (Status: {Status})", 
                            fileItem.FileName, fileItem.ValidationStatus);
                    }

                    UpdateUI();

                    // Start company name scan for the first valid file if in Save Quotes mode
                    if (SaveQuotesMode && string.IsNullOrWhiteSpace(CompanyNameInput) && validatedFiles.Any())
                    {
                        var firstValidFile = validatedFiles.First();
                        StartCompanyNameScan(firstValidFile.FilePath);
                    }
                });

                _logger.Information("Added {ValidCount} valid files out of {TotalCount} dropped files", 
                    validatedFiles.Count, filePaths.Length);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to add files");
                
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    MessageBox.Show($"Failed to add files: {ex.Message}", 
                        "File Addition Error", 
                        MessageBoxButton.OK, 
                        MessageBoxImage.Warning);
                });
            }
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
        
        /// <summary>
        /// Gets the effective company name for processing (user input or detected)
        /// </summary>
        private string GetEffectiveCompanyName()
        {
            // Use typed value first, then detected value
            var companyName = !string.IsNullOrWhiteSpace(CompanyNameInput) 
                ? CompanyNameInput.Trim() 
                : DetectedCompanyName?.Trim();
            
            return companyName ?? string.Empty;
        }
        
        [RelayCommand]
        private async Task ProcessFiles()
        {
            try
            {
                // Handle Save Quotes mode
                if (SaveQuotesMode)
                {
                    await ProcessSaveQuotes();
                    return;
                }

                // Validate and prepare files using DocumentWorkflowService
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
                }
                
                // Get valid files
                var validFiles = PendingFiles.Where(f => 
                    f.ValidationStatus == ValidationStatus.Valid).ToList();

                if (!validFiles.Any())
                {
                    await _uiStateService.UpdateStatusAsync("No valid files to process");
                    return;
                }

                // Set processing state using UIStateService
                await _uiStateService.SetProcessingAsync(true);
                await _uiStateService.UpdateStatusAsync(validFiles.Count > 1 ? "Merging and processing files..." : "Processing file...");

                // Use DocumentWorkflowService for batch processing
                var requests = validFiles.Select(f => new DocumentProcessingRequest
                {
                    FilePath = f.FilePath,
                    OutputPath = _configService.Config.DefaultSaveLocation,
                    Options = new Dictionary<string, object>
                    {
                        ["ConvertOfficeToPdf"] = ConvertOfficeToPdf,
                        ["OpenFolderAfterProcessing"] = OpenFolderAfterProcessing
                    }
                }).ToList();

                var progress = new Progress<BatchProgress>(p =>
                {
                    Application.Current.Dispatcher.InvokeAsync(async () =>
                    {
                        await _uiStateService.UpdateProgressAsync(p.PercentComplete, $"Processing {p.CurrentItem}... ({p.CompletedItems}/{p.TotalItems})");
                    });
                });

                var batchResult = await _documentWorkflowService.ProcessDocumentsAsync(
                    requests, 
                    new BatchProcessingOptions { MaxConcurrency = 3 }, 
                    progress);

                                // Handle results
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (batchResult.SuccessfulFiles > 0)
                    {
                        await _uiStateService.UpdateStatusAsync($"Processing completed: {batchResult.SuccessfulFiles}/{batchResult.TotalFiles} files successful");
                        
                        if (batchResult.SuccessfulFiles == batchResult.TotalFiles)
                        {
                            PendingFiles.Clear();
                            // Success animation could be added here in future
                        }
                    }
                    else
                    {
                                                    await _uiStateService.UpdateStatusAsync("Processing failed - check logs for details");
                    }

                    if (batchResult.FailedFiles > 0)
                    {
                        var failedFiles = batchResult.Results.Where(r => !r.Success).ToList();
                        var errorMessage = $"{batchResult.FailedFiles} files failed to process:\n\n" +
                            string.Join("\n", failedFiles.Select(f => $"• {Path.GetFileName(f.FilePath)}: {f.ErrorMessage}"));
                        
                                                await _uiStateService.ShowWarningAsync("Processing Errors", errorMessage);
                    }
                });
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Unexpected error in ProcessFiles command");
                
                                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    await _uiStateService.UpdateStatusAsync("Processing failed");
                    await _uiStateService.ShowErrorAsync("Error", $"An unexpected error occurred: {ex.Message}");
                });
            }
                                finally
                    {
                        await Application.Current.Dispatcher.InvokeAsync(async () =>
                        {
                            await _uiStateService.SetProcessingAsync(false);
                            await _uiStateService.ResetProgressAsync();
                            UpdateUI();
                        });
                    }
        }
        
        [RelayCommand]
        private async Task ClearFiles()
        {
            PendingFiles.Clear();
            CleanupTempFiles();
            
            // Reset company name detection
            CompanyNameInput = "";
            DetectedCompanyName = "";
            
            await _uiStateService.UpdateStatusAsync("Files cleared");
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
        private void ClearScopeSearch()
        {
            // Clear both search text and selected scope
            ScopeSearchText = "";
            SelectedScope = null;
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
            var window = new EditCompanyNamesWindow((CompanyNameService)_companyNameService)
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
            var window = new Views.EditScopesOfWorkWindow((ScopeOfWorkService)_scopeOfWorkService)
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
                await _uiStateService.UpdateStatusAsync("Running queue diagnostic...");
                
                var diagnosticResult = await Task.Run(async () => 
                {
                    return await QuickDiagnostic.RunQueueDiagnosticAsync();
                });
                
                // Show results in a message box
                MessageBox.Show(diagnosticResult, "Queue Processing Diagnostic Results", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                await _uiStateService.UpdateStatusAsync("Diagnostic completed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during queue diagnostic");
                MessageBox.Show($"Diagnostic failed: {ex.Message}", "Diagnostic Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                await _uiStateService.UpdateStatusAsync("Diagnostic failed");
            }
        }

        [RelayCommand]
        private void OpenSettings()
        {
            var settingsViewModel = new SettingsViewModel(
                _configService, 
                (CompanyNameService)_companyNameService, 
                (ScopeOfWorkService)_scopeOfWorkService);
                
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
                            await _uiStateService.UpdateStatusAsync("Testing company detection...");
            
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
            Application.Current.Dispatcher.Invoke(async () =>
            {
                if (e.Success)
                {
                    await _uiStateService.UpdateQueueStatusAsync("Successfully saved!");
                    
                    // Reset message after 1 second
                    var timer = new DispatcherTimer 
                    { 
                        Interval = TimeSpan.FromSeconds(1) 
                    };
                    timer.Tick += async (s, args) =>
                    {
                        timer.Stop();
                        await UpdateQueueStatusMessage();
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

        private async Task UpdateQueueStatusMessage()
        {
            if (IsQueueProcessing)
            {
                await _uiStateService.UpdateQueueStatusAsync("Processing queue...");
            }
            else if (QueueTotalCount > 0)
            {
                await _uiStateService.UpdateQueueStatusAsync($"{QueueTotalCount} item(s) in queue");
            }
            else
            {
                await _uiStateService.UpdateQueueStatusAsync("Drop quote documents");
            }
        }

        private async Task ProcessSaveQuotes()
        {
            try
            {
                // Check configuration for processing mode
                var saveQuotesConfig = _configService.Config.SaveQuotes;
                
                if (saveQuotesConfig.DefaultProcessingMode == ProcessingMode.Pipeline ||
                    saveQuotesConfig.DefaultProcessingMode == ProcessingMode.Hybrid)
                {
                    // Use the secure pipeline architecture
                    await ProcessSaveQuotesWithPipeline();
                }
                else
                {
                    // Legacy queue processing (for backward compatibility only)
                    await ProcessSaveQuotesLegacy();
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to process save quotes");
                await _uiStateService.UpdateStatusAsync("Processing failed");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// Process save quotes using the secure pipeline architecture
        /// </summary>
        private async Task ProcessSaveQuotesWithPipeline()
        {
            try
            {
                await _uiStateService.SetProcessingAsync(true);
                
                // Validate required fields
                if (string.IsNullOrWhiteSpace(SelectedScope))
                {
                    await _uiStateService.UpdateStatusAsync("Please select a scope of work");
                    MessageBox.Show("Please select a scope of work before processing.", 
                        "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // Determine company name
                var companyName = GetEffectiveCompanyName();
                if (string.IsNullOrWhiteSpace(companyName))
                {
                    await _uiStateService.UpdateStatusAsync("Please enter a company name");
                    MessageBox.Show("Please enter a company name before processing.", 
                        "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // Get valid files
                var validFiles = PendingFiles.Where(f => f.ValidationStatus == ValidationStatus.Valid).ToList();
                if (!validFiles.Any())
                {
                    await _uiStateService.UpdateStatusAsync("No valid files to process");
                    MessageBox.Show("No valid files to process. Please add files first.", 
                        "No Files", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                // Check mode manager availability
                if (_modeManager == null)
                {
                    _logger.Error("Mode manager not available for pipeline processing");
                    await _uiStateService.UpdateStatusAsync("Pipeline processing unavailable");
                    
                    // Fallback to queue if enabled
                    if (_configService.Config.SaveQuotes.EnableQueueFallback)
                    {
                        _logger.Warning("Falling back to queue processing");
                        await ProcessSaveQuotesLegacy();
                        return;
                    }
                    
                    MessageBox.Show("Pipeline processing is not available. Please restart the application.", 
                        "Service Unavailable", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                _logger.Information("Processing {FileCount} files using SaveQuotes pipeline", validFiles.Count);
                
                // Create processing request with pipeline flag
                var request = new ProcessingRequest
                {
                    Files = validFiles,
                    OutputDirectory = SessionSaveLocation ?? _configService.Config.DefaultSaveLocation,
                    Parameters = new Dictionary<string, object>
                    {
                        ["scope"] = SelectedScope,
                        ["companyName"] = companyName,
                        ["usePipeline"] = true,  // CRITICAL: Enable pipeline processing
                        ["enableSecurityValidation"] = _configService.Config.SaveQuotes.EnableSecurityValidation,
                        ["maxRetryAttempts"] = _configService.Config.SaveQuotes.MaxRetryAttempts
                    }
                };
                
                // Process using mode system with pipeline
                var result = await _modeManager.ProcessFilesAsync(
                    request.Files, 
                    request.OutputDirectory, 
                    request.Parameters);
                
                // Handle results
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (result.Success)
                    {
                        await _uiStateService.UpdateStatusAsync($"Successfully processed {result.ProcessedFiles.Count} files");
                        
                        // Clear UI
                        PendingFiles.Clear();
                        CompanyNameInput = "";
                        DetectedCompanyName = "";
                        
                        // Clear scope if configured
                        if (_configService.Config.ClearScopeAfterProcessing)
                        {
                            ScopeSearchText = "";
                            SelectedScope = null;
                        }
                        
                        // Open output folder if configured
                        if (_configService.Config.OpenFolderAfterProcessing == true && !string.IsNullOrEmpty(request.OutputDirectory))
                        {
                            OpenOutputFolder(request.OutputDirectory);
                        }
                        
                        _logger.Information("Pipeline processing completed successfully");
                    }
                    else
                    {
                        await _uiStateService.UpdateStatusAsync($"Processing failed: {result.ErrorMessage}");
                        
                        var failedCount = result.ProcessedFiles.Count(f => !f.Success);
                        var message = $"Processing completed with errors:\n\n" +
                                    $"Total files: {result.ProcessedFiles.Count}\n" +
                                    $"Failed: {failedCount}\n\n" +
                                    $"Error: {result.ErrorMessage}";
                        
                        MessageBox.Show(message, "Processing Errors", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        
                        _logger.Error("Pipeline processing failed: {Error}", result.ErrorMessage);
                    }
                });
            }
            finally
            {
                await _uiStateService.SetProcessingAsync(false);
                UpdateUI();
            }
        }
        
        /// <summary>
        /// Legacy queue-based processing (for backward compatibility)
        /// </summary>
        private async Task ProcessSaveQuotesLegacy()
        {
            try
            {
                await _uiStateService.SetProcessingAsync(true);
                
                // Validate required fields
                if (string.IsNullOrWhiteSpace(SelectedScope))
                {
                    await _uiStateService.UpdateStatusAsync("Please select a scope of work");
                    MessageBox.Show("Please select a scope of work before processing.", 
                        "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // Determine company name
                var companyName = GetEffectiveCompanyName();
                if (string.IsNullOrWhiteSpace(companyName))
                {
                    await _uiStateService.UpdateStatusAsync("Please enter a company name");
                    MessageBox.Show("Please enter a company name before processing.", 
                        "Missing Information", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // Get valid files
                var validFiles = PendingFiles.Where(f => f.ValidationStatus == ValidationStatus.Valid).ToList();
                if (!validFiles.Any())
                {
                    await _uiStateService.UpdateStatusAsync("No valid files to process");
                    MessageBox.Show("No valid files to process. Please add files first.", 
                        "No Files", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                // Check mode manager availability
                if (_modeManager == null)
                {
                    _logger.Error("Mode manager not available for queue processing");
                    await _uiStateService.UpdateStatusAsync("Queue processing unavailable");
                    
                    // Fallback to direct processing if no other option
                    _logger.Warning("Queue processing unavailable with no fallback");
                    
                    MessageBox.Show("Queue processing is not available. Please restart the application.", 
                        "Service Unavailable", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                _logger.Information("Processing {FileCount} files using SaveQuotes queue", validFiles.Count);
                
                // Create processing request with queue flag
                var request = new ProcessingRequest
                {
                    Files = validFiles,
                    OutputDirectory = SessionSaveLocation ?? _configService.Config.DefaultSaveLocation,
                    Parameters = new Dictionary<string, object>
                    {
                        ["scope"] = SelectedScope,
                        ["companyName"] = companyName,
                        ["useQueue"] = true,  // CRITICAL: Enable queue processing
                        ["enableSecurityValidation"] = _configService.Config.SaveQuotes.EnableSecurityValidation,
                        ["maxRetryAttempts"] = _configService.Config.SaveQuotes.MaxRetryAttempts
                    }
                };
                
                // Process using mode system with queue
                var result = await _modeManager.ProcessFilesAsync(
                    request.Files, 
                    request.OutputDirectory, 
                    request.Parameters);
                
                // Handle results
                await Application.Current.Dispatcher.InvokeAsync(async () =>
                {
                    if (result.Success)
                    {
                        await _uiStateService.UpdateStatusAsync($"Successfully processed {result.ProcessedFiles.Count} files");
                        
                        // Clear UI
                        PendingFiles.Clear();
                        CompanyNameInput = "";
                        DetectedCompanyName = "";
                        
                        // Clear scope if configured
                        if (_configService.Config.ClearScopeAfterProcessing)
                        {
                            ScopeSearchText = "";
                            SelectedScope = null;
                        }
                        
                        // Open output folder if configured
                        if (_configService.Config.OpenFolderAfterProcessing == true && !string.IsNullOrEmpty(request.OutputDirectory))
                        {
                            OpenOutputFolder(request.OutputDirectory);
                        }
                        
                        _logger.Information("Queue processing completed successfully");
                    }
                    else
                    {
                        await _uiStateService.UpdateStatusAsync($"Processing failed: {result.ErrorMessage}");
                        
                        var failedCount = result.ProcessedFiles.Count(f => !f.Success);
                        var message = $"Processing completed with errors:\n\n" +
                                    $"Total files: {result.ProcessedFiles.Count}\n" +
                                    $"Failed: {failedCount}\n\n" +
                                    $"Error: {result.ErrorMessage}";
                        
                        MessageBox.Show(message, "Processing Errors", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                        
                        _logger.Error("Queue processing failed: {Error}", result.ErrorMessage);
                    }
                });
            }
            finally
            {
                await _uiStateService.SetProcessingAsync(false);
                UpdateUI();
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
        
        // Add disposal guard flag
        private volatile bool _isCleaningUp = false;
        
        public void Cleanup()
        {
            try
            {
                _logger.Information("Starting MainViewModel cleanup");
                
                // Stop queue processing first
                if (_queueService != null)
                {
                    try
                    {
                        _queueService.StopProcessing();
                        _logger.Information("Queue processing stopped");
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error stopping queue processing");
                    }
                }
                
                // Cancel any pending scans
                try
                {
                    _currentScanCancellation?.Cancel();
                    _currentScanCancellation?.Dispose();
                    _currentScanCancellation = null;
                    _logger.Information("Cancelled pending company name scans");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error cancelling company name scans");
                }
                
                // Dispose services in dependency order
                try
                {
                    // Company name service (depends on Office services)
                    _companyNameService?.Dispose();
                    _logger.Information("Company name service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing company name service");
                }
                
                try
                {
                    // File processing service
                    _fileProcessingService?.Dispose();
                    _logger.Information("File processing service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing file processing service");
                }
                
                try
                {
                    // Queue service
                    _queueService?.Dispose();
                    _logger.Information("Queue service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing queue service");
                }
                
                // MEMORY FIX: SessionAware services are now null, so no disposal needed
                // if (_sessionOfficeService != null)
                // {
                //     try
                //     {
                //         _sessionOfficeService.Dispose();
                //         _logger.Information("Session Office service disposed");
                //     }
                //     catch (Exception ex)
                //     {
                //         _logger.Warning(ex, "Error disposing session Office service");
                //     }
                // }
                
                // if (_sessionExcelService != null)
                // {
                //     try
                //     {
                //         _sessionExcelService.Dispose();
                //         _logger.Information("Session Excel service disposed");
                //     }
                //     catch (Exception ex)
                //     {
                //         _logger.Warning(ex, "Error disposing session Excel service");
                //     }
                // }
                
                // Regular Office conversion service
                try
                {
                    _officeConversionService?.Dispose();
                    _logger.Information("Office conversion service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing office conversion service");
                }
                
                // Health monitor
                try
                {
                    _healthMonitor?.Dispose();
                    _logger.Information("Health monitor disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing health monitor");
                }
                
                // Performance monitor
                try
                {
                    _performanceMonitor?.Dispose();
                    _logger.Information("Performance monitor disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing performance monitor");
                }
                
                // Process manager
                try
                {
                    _processManager?.Dispose();
                    _logger.Information("Process manager disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing process manager");
                }
                
                // PDF cache service
                try
                {
                    _pdfCacheService?.Dispose();
                    _logger.Information("PDF cache service disposed");
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error disposing PDF cache service");
                }
                
                // Force COM cleanup and garbage collection
                ComHelper.ForceComCleanup("MainViewModelCleanup");
                
                // Wait for any remaining cleanup
                Thread.Sleep(1000);
                
                // Final COM statistics
                var finalStats = ComHelper.GetComObjectSummary();
                _logger.Information("Final cleanup complete - COM Objects: Created {Created}, Released {Released}, Net {Net}", 
                    finalStats.TotalCreated, finalStats.TotalReleased, finalStats.NetObjects);
                
                // Clean up any temporary files
                CleanupTempFiles();
                
                _logger.Information("MainViewModel cleanup completed successfully");
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
        private bool _isDisposing = false;
        
        public void Dispose()
        {
            // Check _disposed first
            if (_disposed)
            {
                _logger.Warning("MainViewModel.Dispose called on already disposed object");
                return;
            }
            
            if (_isDisposing) 
            {
                _logger.Warning("MainViewModel.Dispose already in progress");
                return;
            }
            _isDisposing = true;
            
            _logger.Information("MainViewModel.Dispose starting...");
            
            Dispose(true);
            GC.SuppressFinalize(this);
            
            // Force COM cleanup and wait briefly for it to complete
            ComHelper.ForceComCleanup("MainViewModelDispose");
            System.Threading.Thread.Sleep(100); // Give COM time to clean up

            // Log final COM statistics to verify cleanup
            var summary = ComHelper.GetComObjectSummary();
            _logger.Information("Final COM statistics after disposal - Net Objects: {NetObjects}", summary.NetObjects);

            if (summary.NetObjects > 0)
            {
                _logger.Error("COM objects still not released after disposal! Check disposal chain.");
                ComHelper.LogComObjectStats();
            }

            _logger.Information("MainViewModel disposed");
        }
        
        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Cancel all operations first
                    try
                    {
                        _viewModelCts?.Cancel();
                        _filterCts?.Cancel();
                        _currentScanCancellation?.Cancel();
                        
                        // Wait a bit for operations to cancel
                        Task.Delay(100).Wait();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error cancelling operations during disposal");
                    }
                    
                    // Unsubscribe from all events to prevent memory leaks
                    UnsubscribeFromEvents();
                    
                    // Dispose cancellation tokens
                    try
                    {
                        _viewModelCts?.Dispose();
                        _filterCts?.Dispose();
                        _currentScanCancellation?.Dispose();
                        _scanSemaphore?.Dispose();
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Error disposing cancellation tokens");
                    }
                    
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

        // Advanced Mode UI Framework Commands (Phase 2 Milestone 2 - Day 5)
        [RelayCommand]
        private async Task SwitchToStandardModeAsync()
        {
            try
            {
                await _advancedModeUIManager.SwitchToModeAsync("default");
                await _uiStateService.UpdateStatusAsync("Switched to Standard Processing mode");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error switching to standard mode");
                StatusMessage = "Error switching mode";
            }
        }

        [RelayCommand]
        private async Task SwitchToSaveQuotesModeAsync()
        {
            try
            {
                await _advancedModeUIManager.SwitchToModeAsync("SaveQuotes");
                await _uiStateService.UpdateStatusAsync("Switched to Save Quotes mode");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error switching to save quotes mode");
                StatusMessage = "Error switching mode";
            }
        }

        [RelayCommand]
        private async Task ToggleAdvancedModeUIAsync()
        {
            try
            {
                IsAdvancedModeUIEnabled = !IsAdvancedModeUIEnabled;
                StatusMessage = IsAdvancedModeUIEnabled ? "Advanced Mode UI enabled" : "Advanced Mode UI disabled";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error toggling advanced mode UI");
                StatusMessage = "Error toggling mode UI";
            }
        }

        // Manual command property - workaround for source generator issue
        private IRelayCommand? _testMemoryLeakFixesCommand;
        public IRelayCommand TestMemoryLeakFixesCommand => 
            _testMemoryLeakFixesCommand ??= new AsyncRelayCommand(TestMemoryLeakFixesAsync);
        
        private async Task TestMemoryLeakFixesAsync()
        {
            try
            {
                await _uiStateService.UpdateStatusAsync("Running memory leak test...");
                
                var testResult = await Task.Run(async () => 
                {
                    return await QuickDiagnostic.TestMemoryLeakFixes();
                });
                
                // Show results in a message box
                MessageBox.Show(testResult, "Memory Leak Test Results", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                await _uiStateService.UpdateStatusAsync("Memory leak test completed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during memory leak test");
                MessageBox.Show($"Memory leak test failed: {ex.Message}", "Test Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusMessage = "Memory leak test failed";
            }
        }

        [RelayCommand]
        private async Task TestThreadSafetyAsync()
        {
            try
            {
                StatusMessage = "Running thread safety test...";
                
                var testResult = await Task.Run(async () => 
                {
                    return await QuickDiagnostic.TestThreadSafetyImprovements();
                });
                
                // Show results in a message box
                MessageBox.Show(testResult, "Thread Safety Test Results", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                await _uiStateService.UpdateStatusAsync("Thread safety test completed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during thread safety test");
                MessageBox.Show($"Thread safety test failed: {ex.Message}", "Test Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusMessage = "Thread safety test failed";
            }
        }

        [RelayCommand]
        private async Task TestErrorRecoveryAsync()
        {
            try
            {
                StatusMessage = "Running error recovery test...";
                
                var testResult = await Task.Run(async () => 
                {
                    return await QuickDiagnostic.TestErrorRecoveryImprovements();
                });
                
                // Show results in a message box
                MessageBox.Show(testResult, "Error Recovery Test Results", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                StatusMessage = "Error recovery test completed";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during error recovery test");
                
                // Use our error recovery service if available
                if (_errorRecoveryService != null)
                {
                    var errorInfo = _errorRecoveryService.CreateErrorInfo(ex, "Error recovery test");
                    MessageBox.Show(
                        $"{errorInfo.Message}\n\n{errorInfo.RecoveryGuidance}",
                        errorInfo.Title,
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show($"Error recovery test failed: {ex.Message}", "Test Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
                StatusMessage = "Error recovery test failed";
            }
        }

        [RelayCommand]
        private async Task TestModeSystemAsync()
        {
            try
            {
                StatusMessage = "Running mode system test...";
                
                var testResult = await Task.Run(async () => 
                {
                    return await QuickDiagnostic.TestModeSystemInfrastructure();
                });
                
                // Show results in a message box
                MessageBox.Show(testResult, "Mode System Test Results", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
                
                StatusMessage = "Mode system test completed";
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during mode system test");
                
                // Use our error recovery service if available
                if (_errorRecoveryService != null)
                {
                    var errorInfo = _errorRecoveryService.CreateErrorInfo(ex, "Mode system test");
                    MessageBox.Show(
                        $"{errorInfo.Message}\n\n{errorInfo.RecoveryGuidance}",
                        errorInfo.Title,
                        MessageBoxButton.OK,
                        MessageBoxImage.Error);
                }
                else
                {
                    MessageBox.Show($"Mode system test failed: {ex.Message}", "Test Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
                
                StatusMessage = "Mode system test failed";
            }
        }
        
        // Command property auto-generated by CommunityToolkit.Mvvm from [RelayCommand] attribute
    }
}
