using System.Windows;

namespace DocHandler.Views
{
    public partial class SimpleInputDialog : Window
    {
        public string InputText { get; private set; } = "";
        
        public SimpleInputDialog(string prompt, string title, string defaultValue = "")
        {
            InitializeComponent();
            
            Title = title;
            PromptText.Text = prompt;
            InputBox.Text = defaultValue;
            
            Loaded += (s, e) =>
            {
                InputBox.Focus();
                InputBox.SelectAll();
            };
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            InputText = InputBox.Text;
            DialogResult = true;
        }
        
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}