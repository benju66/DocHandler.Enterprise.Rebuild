using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace DocHandler.Views
{
    public partial class ScopeEditDialog : Window, INotifyPropertyChanged
    {
        private string _scopeCode = string.Empty;
        private string _scopeDescription = string.Empty;
        
        public string ScopeCode 
        { 
            get => _scopeCode;
            set
            {
                if (_scopeCode != value)
                {
                    _scopeCode = value ?? string.Empty;
                    OnPropertyChanged();
                }
            }
        }
        
        public string ScopeDescription 
        { 
            get => _scopeDescription;
            set
            {
                if (_scopeDescription != value)
                {
                    _scopeDescription = value ?? string.Empty;
                    OnPropertyChanged();
                }
            }
        }
        
        public event PropertyChangedEventHandler? PropertyChanged;
        
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        
        public ScopeEditDialog(string scopeCode, string scopeDescription)
        {
            InitializeComponent();
            
            ScopeCode = scopeCode ?? string.Empty;
            ScopeDescription = scopeDescription ?? string.Empty;
            
            DataContext = this;
            
            // Focus on scope code box
            Loaded += (s, e) => ScopeCodeBox.Focus();
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            ScopeCode = ScopeCode?.Trim() ?? string.Empty;
            ScopeDescription = ScopeDescription?.Trim() ?? string.Empty;
            
            if (string.IsNullOrWhiteSpace(ScopeCode) || string.IsNullOrWhiteSpace(ScopeDescription))
            {
                MessageBox.Show("Both scope code and description are required.", "Validation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            
            DialogResult = true;
        }
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
} 