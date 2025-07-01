using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
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
        
        public EditCompanyNamesViewModel(CompanyNameService companyNameService)
        {
            _logger = Log.ForContext<EditCompanyNamesViewModel>();
            _companyNameService = companyNameService;
            
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
            
            FilterCompanies();
        }
        
        partial void OnSearchTextChanged(string value)
        {
            FilterCompanies();
        }
        
        private void FilterCompanies()
        {
            FilteredCompanies.Clear();
            
            var searchTerm = SearchText?.Trim() ?? "";
            var filtered = string.IsNullOrWhiteSpace(searchTerm)
                ? Companies
                : Companies.Where(c => 
                    c.Name.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                    c.AliasesDisplay.Contains(searchTerm, StringComparison.OrdinalIgnoreCase));
            
            foreach (var company in filtered.OrderBy(c => c.Name))
            {
                FilteredCompanies.Add(company);
            }
        }
        
        [RelayCommand]
        private async Task AddCompany()
        {
            var dialog = new Views.CompanyEditDialog("", new())
            {
                Owner = Application.Current.Windows.OfType<Window>().FirstOrDefault(w => w.IsActive),
                Title = "Add New Company"
            };
            
            if (dialog.ShowDialog() == true)
            {
                var success = await _companyNameService.AddCompanyName(dialog.CompanyName, dialog.Aliases);
                
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
                        Aliases = dialog.Aliases,
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
                    company.CompanyInfo.Aliases = dialog.Aliases;
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
                    Companies.Remove(company);
                    FilteredCompanies.Remove(company);
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
        private async Task RefreshList()
        {
            await Task.Run(() => LoadCompanies());
        }
        
        public void MarkAsChanged()
        {
            HasChanges = true;
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