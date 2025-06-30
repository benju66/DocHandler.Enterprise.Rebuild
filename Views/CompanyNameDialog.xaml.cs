using System.Collections.ObjectModel;
using System.Windows;

namespace DocHandler.Views
{
    public partial class CompanyNameDialog : Window
    {
        public string CompanyName { get; set; }
        public bool AddToDatabase { get; set; } = true;
        public string Message { get; set; }
        public ObservableCollection<string> SuggestedCompanies { get; set; }

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