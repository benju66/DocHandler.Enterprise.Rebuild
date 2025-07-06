using System.Windows;
using DocHandler.Services;
using DocHandler.ViewModels;
using Serilog;

namespace DocHandler.Views
{
    public partial class EditScopesOfWorkWindow : Window
    {
        private readonly EditScopesOfWorkViewModel _viewModel;
        
        public EditScopesOfWorkWindow(ScopeOfWorkService scopeService)
        {
            InitializeComponent();
            var logger = Log.ForContext<EditScopesOfWorkViewModel>();
            _viewModel = new EditScopesOfWorkViewModel(scopeService, logger);
            DataContext = _viewModel;
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = _viewModel.HasChanges;
            Close();
        }
    }
} 