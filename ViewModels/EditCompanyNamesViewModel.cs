using System;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using Serilog;

namespace DocHandler.ViewModels
{
    public partial class EditCompanyNamesViewModel : ObservableObject
    {
        private readonly ILogger _logger;
        private readonly CompanyNameService _companyNameService;
        private DispatcherTimer _searchTimer;
        private CancellationTokenSource _searchCancellation;
        private const int SearchDelayMs = 300; // Wait 300ms after user stops typing
        
        [ObservableProperty]
        private ObservableCollection<CompanyItemViewModel> _companies = new();
        
        [ObservableProperty]
        private ObservableCollection<CompanyItemViewModel> _filteredCompanies = new();
        
        [ObservableProperty]
        private CompanyItemViewModel? _selectedCompany;
        
        [ObservableProperty]
        private string _searchText = "";
        
        [ObservableProperty]
        private bool _hasChanges;
        
        public int CompanyCount => Companies.Count;
        public int FilteredCompanyCount => FilteredCompanies.Count;
        
        public EditCompanyNamesViewModel(CompanyNameService companyNameService)
        {
            _logger = Log.ForContext<EditCompanyNamesViewModel>();
            _companyNameService = companyNameService;
            
            // Initialize search timer for debouncing
            _searchTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(SearchDelayMs)
            };
            _searchTimer.Tick += async (s, e) =>
            {
                _searchTimer.Stop();
                await FilterCompaniesAsync();
            };
            
            LoadCompanies();
        }
        
        private void LoadCompanies()
        {
            Companies.Clear();
            
            foreach (var company in _companyNameService.Companies.OrderBy(c => c.Name))
            {
                var vm = new CompanyItemViewModel(company)
                {
                    Parent = this
                };
                Companies.Add(vm);
            }
            
            _ = FilterCompaniesAsync();
            OnPropertyChanged(nameof(CompanyCount));
        }
        
        partial void OnSearchTextChanged(string value)
        {
            // Cancel previous search
            _searchCancellation?.Cancel();
            _searchCancellation = new CancellationTokenSource();
            
            // Restart the timer for debouncing
            _searchTimer.Stop();
            _searchTimer.Start();
        }
        
        private async Task FilterCompaniesAsync()
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
                        return Companies.ToList();
                    }
                    
                    return Companies.Where(c => 
                        c.Name.Contains(searchText, StringComparison.OrdinalIgnoreCase) ||
                        c.AliasesDisplay.Contains(searchText, StringComparison.OrdinalIgnoreCase))
                        .OrderBy(c => c.Name)
                        .ToList();
                }, cancellationToken);
                
                // Update UI on the UI thread
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (!cancellationToken.IsCancellationRequested)
                    {
                        FilteredCompanies.Clear();
                        foreach (var company in filteredList)
                        {
                            FilteredCompanies.Add(company);
                        }
                        OnPropertyChanged(nameof(FilteredCompanyCount));
                    }
                });
            }
            catch (OperationCanceledException)
            {
                // Search was cancelled, this is expected
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to filter companies");
            }
        }
        
        [RelayCommand]
        private async Task AddCompany()
        {
            var dialog = new Views.CompanyEditDialog("", new List<string>())
            {
                Owner = Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive),
                Title = "Add New Company"
            };
            
            if (dialog.ShowDialog() == true)
            {
                var success = await _companyNameService.AddCompanyName(dialog.CompanyName, dialog.Aliases.ToList());
                
                if (success)
                {
                    LoadCompanies();
                    HasChanges = true;
                    _logger.Information("Added new company: {Name}", dialog.CompanyName);
                }
                else
                {
                    MessageBox.Show("Failed to add company. It may already exist.", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        
        [RelayCommand]
        private async Task EditCompany(CompanyItemViewModel? company)
        {
            if (company == null) return;
            
            var dialog = new Views.CompanyEditDialog(company.Name, company.Aliases)
            {
                Owner = Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive),
                Title = "Edit Company"
            };
            
            if (dialog.ShowDialog() == true)
            {
                // If name changed, we need to remove old and add new
                if (!company.Name.Equals(dialog.CompanyName, StringComparison.OrdinalIgnoreCase))
                {
                    // Remove old
                    await _companyNameService.RemoveCompanyName(company.Name);
                    
                    // Add new with preserved usage stats
                    var newCompany = new CompanyInfo
                    {
                        Name = dialog.CompanyName,
                        Aliases = dialog.Aliases.ToList(),
                        DateAdded = company.CompanyInfo.DateAdded,
                        LastUsed = company.CompanyInfo.LastUsed,
                        UsageCount = company.CompanyInfo.UsageCount
                    };
                    
                    _companyNameService.Companies.Add(newCompany);
                    await _companyNameService.SaveCompanyNames();
                }
                else
                {
                    // Just update aliases
                    company.CompanyInfo.Aliases = dialog.Aliases.ToList();
                    await _companyNameService.SaveCompanyNames();
                }
                
                LoadCompanies();
                HasChanges = true;
                _logger.Information("Updated company: {OldName} -> {NewName}", company.Name, dialog.CompanyName);
            }
        }
        
        [RelayCommand]
        private async Task DeleteCompany(CompanyItemViewModel? company)
        {
            if (company == null) return;
            
            var result = MessageBox.Show(
                $"Are you sure you want to delete '{company.Name}'?\n\nThis action cannot be undone.",
                "Confirm Delete",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            
            if (result == MessageBoxResult.Yes)
            {
                var success = await _companyNameService.RemoveCompanyName(company.Name);
                
                if (success)
                {
                    LoadCompanies();
                    HasChanges = true;
                    _logger.Information("Deleted company: {Name}", company.Name);
                }
            }
        }
        
        [RelayCommand]
        private void ClearSearch()
        {
            SearchText = "";
        }
        
        [RelayCommand]
        private void RefreshList()
        {
            try
            {
                LoadCompanies();
                _logger.Information("Company list refreshed successfully");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to refresh company list");
                MessageBox.Show("Failed to refresh list. Please try again.", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        public void MarkAsChanged()
        {
            HasChanges = true;
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
    
    public partial class CompanyItemViewModel : ObservableObject
    {
        public CompanyInfo CompanyInfo { get; }
        
        public string Name => CompanyInfo.Name;
        public List<string> Aliases => CompanyInfo.Aliases;
        public DateTime DateAdded => CompanyInfo.DateAdded;
        public DateTime? LastUsed => CompanyInfo.LastUsed;
        public int UsageCount => CompanyInfo.UsageCount;
        
        public string AliasesDisplay => Aliases.Any() ? string.Join(", ", Aliases) : "None";
        public string LastUsedDisplay => LastUsed?.ToString("MMM d, yyyy") ?? "Never";
        public string UsageDisplay => UsageCount == 1 ? "1 time" : $"{UsageCount} times";
        
        public EditCompanyNamesViewModel? Parent { get; set; }
        
        public CompanyItemViewModel(CompanyInfo companyInfo)
        {
            CompanyInfo = companyInfo;
        }
    }
}