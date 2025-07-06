using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using DocHandler.Views;
using Serilog;

namespace DocHandler.ViewModels
{
    public partial class EditScopesOfWorkViewModel : ObservableObject
    {
        private readonly ScopeOfWorkService _scopeService;
        private readonly ILogger _logger;
        private string _searchText = string.Empty;
        private DispatcherTimer _searchTimer;
        private CancellationTokenSource _searchCancellation;
        private const int SearchDelayMs = 300; // Wait 300ms after user stops typing

        public ObservableCollection<ScopeOfWork> Scopes { get; } = new();
        public ObservableCollection<ScopeOfWork> FilteredScopes { get; } = new();

        public string SearchText
        {
            get => _searchText;
            set
            {
                if (SetProperty(ref _searchText, value))
                {
                    // Cancel previous search
                    _searchCancellation?.Cancel();
                    _searchCancellation = new CancellationTokenSource();
                    
                    // Restart the timer for debouncing
                    _searchTimer.Stop();
                    _searchTimer.Start();
                }
            }
        }

        [ObservableProperty]
        private bool _hasChanges;

        public int ScopeCount => Scopes.Count;
        public int FilteredScopeCount => FilteredScopes.Count;

        public EditScopesOfWorkViewModel(ScopeOfWorkService scopeService, ILogger logger)
        {
            _scopeService = scopeService;
            _logger = logger;
            
            // Initialize search timer for debouncing
            _searchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(SearchDelayMs)
            };
            _searchTimer.Tick += async (s, e) =>
            {
                _searchTimer.Stop();
                await FilterScopesAsync();
            };
            
            LoadScopes();
        }

        private void LoadScopes()
        {
            try
            {
                var scopes = _scopeService.Scopes; // Use the existing Scopes property
                Scopes.Clear();
                foreach (var scope in scopes.OrderBy(s => s.Code))
                {
                    Scopes.Add(scope);
                }
                FilterScopes();
                OnPropertyChanged(nameof(ScopeCount));
                _logger.Information("Loaded {Count} scopes", Scopes.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load scopes");
                MessageBox.Show("Failed to load scopes. Please try again.", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FilterScopes()
        {
            FilteredScopes.Clear();
            
            var filtered = string.IsNullOrWhiteSpace(SearchText)
                ? Scopes
                : Scopes.Where(s => 
                    s.Code.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
                    s.Description.Contains(SearchText, StringComparison.OrdinalIgnoreCase));

            foreach (var scope in filtered)
            {
                FilteredScopes.Add(scope);
            }
            
            OnPropertyChanged(nameof(FilteredScopeCount));
        }

        private async Task FilterScopesAsync()
        {
            try
            {
                var cancellationToken = _searchCancellation?.Token ?? CancellationToken.None;
                
                var searchText = SearchText?.Trim() ?? "";
                var filteredList = await Task.Run(() =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    
                    if (string.IsNullOrWhiteSpace(searchText))
                    {
                        return Scopes.ToList();
                    }
                    
                    return Scopes.Where(s => 
                        s.Code.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                        s.Description.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                }, cancellationToken);
                
                // Update UI on the UI thread
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (!cancellationToken.IsCancellationRequested)
                    {
                        FilteredScopes.Clear();
                        foreach (var scope in filteredList)
                        {
                            FilteredScopes.Add(scope);
                        }
                        OnPropertyChanged(nameof(FilteredScopeCount));
                    }
                });
            }
            catch (OperationCanceledException)
            {
                // Search was cancelled, this is expected
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to filter scopes");
            }
        }

        [RelayCommand]
        private async Task AddScope()
        {
            var dialog = new ScopeEditDialog("", ""); // Provide empty parameters for new scope
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var success = await _scopeService.AddScope(dialog.ScopeCode, dialog.ScopeDescription);
                    if (success)
                    {
                        LoadScopes();
                        HasChanges = true;
                        _logger.Information("Added new scope: {Code}", dialog.ScopeCode);
                    }
                    else
                    {
                        MessageBox.Show("Failed to add scope. The code may already exist.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to add scope");
                    MessageBox.Show("Failed to add scope. Please try again.", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        [RelayCommand]
        private async Task EditScope(ScopeOfWork scope)
        {
            if (scope == null) return;

            var dialog = new ScopeEditDialog(scope.Code, scope.Description);
            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var success = await _scopeService.UpdateScope(scope.Code, dialog.ScopeCode, dialog.ScopeDescription);
                    if (success)
                    {
                        LoadScopes();
                        HasChanges = true;
                        _logger.Information("Updated scope: {Code}", scope.Code);
                    }
                    else
                    {
                        MessageBox.Show("Failed to update scope. The new code may already exist.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to update scope");
                    MessageBox.Show("Failed to update scope. Please try again.", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        [RelayCommand]
        private async Task DeleteScope(ScopeOfWork scope)
        {
            if (scope == null) return;

            var result = MessageBox.Show(
                $"Are you sure you want to delete the scope '{scope.Code}'?",
                "Confirm Delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    var success = await _scopeService.RemoveScope(scope.Code);
                    if (success)
                    {
                        LoadScopes();
                        HasChanges = true;
                        _logger.Information("Deleted scope: {Code}", scope.Code);
                    }
                    else
                    {
                        MessageBox.Show("Failed to delete scope. Please try again.", "Error", 
                            MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to delete scope");
                    MessageBox.Show("Failed to delete scope. Please try again.", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        [RelayCommand]
        private void RefreshList()
        {
            try
            {
                LoadScopes();
                _logger.Information("Scope list refreshed successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to refresh scope list");
                MessageBox.Show("Failed to refresh list. Please try again.", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        [RelayCommand]
        private void ClearSearch()
        {
            SearchText = string.Empty;
        }
        
        public void Dispose()
        {
            _searchTimer?.Stop();
            _searchTimer = null;
            _searchCancellation?.Cancel();
            _searchCancellation?.Dispose();
            _searchCancellation = null;
        }
    }
} 