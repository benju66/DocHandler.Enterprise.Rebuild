using System.ComponentModel;
using System.Windows;
using DocHandler.ViewModels;

namespace DocHandler.Views
{
    public partial class SettingsWindow : Window
    {
        private readonly SettingsViewModel _viewModel;
        
        public SettingsWindow(SettingsViewModel viewModel)
        {
            InitializeComponent();
            _viewModel = viewModel;
            DataContext = _viewModel;
        }
        
        private void Window_Closing(object sender, CancelEventArgs e)
        {
            if (_viewModel.HasUnsavedChanges)
            {
                var result = MessageBox.Show(
                    "You have unsaved changes. Do you want to save them before closing?",
                    "Unsaved Changes",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    if (!_viewModel.SaveSettings())
                    {
                        e.Cancel = true; // Cancel close if save failed
                    }
                }
                else if (result == MessageBoxResult.Cancel)
                {
                    e.Cancel = true; // Cancel close
                }
            }
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (_viewModel.SaveSettings())
            {
                DialogResult = true;
                Close();
            }
        }
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
        
        private void ApplyButton_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.SaveSettings();
        }
    }
} 