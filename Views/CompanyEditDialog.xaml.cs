using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace DocHandler.Views
{
    public partial class CompanyEditDialog : Window
    {
        public string CompanyName { get; set; }
        public ObservableCollection<string> Aliases { get; set; }
        
        public CompanyEditDialog(string companyName, List<string> aliases)
        {
            InitializeComponent();
            
            CompanyName = companyName;
            Aliases = new ObservableCollection<string>(aliases ?? new List<string>());
            
            DataContext = this;
            
            // Focus on company name box
            Loaded += (s, e) => CompanyNameBox.Focus();
        }
        
        private void AddAliasButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SimpleInputDialog(
                "Enter an alias for this company:", 
                "Add Alias")
            {
                Owner = this
            };
            
            if (dialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(dialog.InputText))
            {
                var alias = dialog.InputText.Trim();
                
                // Check if already exists
                if (Aliases.Any(a => a.Equals(alias, System.StringComparison.OrdinalIgnoreCase)))
                {
                    MessageBox.Show("This alias already exists.", "Duplicate Alias", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                Aliases.Add(alias);
            }
        }
        
        private void RemoveAliasButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.Tag is string alias)
            {
                Aliases.Remove(alias);
            }
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            CompanyName = CompanyName.Trim();
            DialogResult = true;
        }
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}