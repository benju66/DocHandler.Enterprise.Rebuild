using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;

namespace DocHandler.Views
{
    public partial class CompanyNameDialog : Window, INotifyPropertyChanged
    {
        private string _companyName = string.Empty;
        
        public string CompanyName 
        { 
            get => _companyName;
            set
            {
                if (_companyName != value)
                {
                    _companyName = value ?? string.Empty;
                    OnPropertyChanged();
                }
            }
        }
        
        public bool AddToDatabase { get; set; } = true;
        public string Message { get; set; }
        public ObservableCollection<string> SuggestedCompanies { get; set; }

        public event PropertyChangedEventHandler? PropertyChanged;
        
        protected virtual void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public CompanyNameDialog(string message, ObservableCollection<string> suggestedCompanies = null)
        {
            InitializeComponent();
            DataContext = this;
            
            Message = message;
            SuggestedCompanies = suggestedCompanies ?? new ObservableCollection<string>();
            
            // Focus on the input box when dialog opens
            Loaded += (s, e) => CompanyNameBox.Focus();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}