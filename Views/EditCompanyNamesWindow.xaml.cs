using System.Windows;
using DocHandler.Services;
using DocHandler.ViewModels;

namespace DocHandler.Views
{
    public partial class EditCompanyNamesWindow : Window
    {
        private readonly EditCompanyNamesViewModel _viewModel;
        
        public EditCompanyNamesWindow(CompanyNameService companyNameService)
        {
            InitializeComponent();
            
            _viewModel = new EditCompanyNamesViewModel(companyNameService);
            DataContext = _viewModel;
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = _viewModel.HasChanges;
            Close();
        }
    }
}