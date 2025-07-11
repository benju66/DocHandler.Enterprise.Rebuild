using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using DocHandler.ViewModels;

namespace DocHandler.Views
{
    public partial class PerformanceMetricsWindow : Window
    {
        private readonly MainViewModel _mainViewModel;
        private readonly DispatcherTimer _refreshTimer;
        private string _currentMetrics;
        
        public PerformanceMetricsWindow(MainViewModel mainViewModel)
        {
            InitializeComponent();
            _mainViewModel = mainViewModel;
            
            _refreshTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(5)
            };
            _refreshTimer.Tick += async (s, e) => await RefreshMetricsAsync();
            
            // Load metrics on startup
            Loaded += async (s, e) => await RefreshMetricsAsync();
        }
        
        private async void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            await RefreshMetricsAsync();
        }
        
        private async Task RefreshMetricsAsync()
        {
            RefreshButton.IsEnabled = false;
            
            try
            {
                _currentMetrics = await _mainViewModel.CollectPerformanceMetricsAsync();
                MetricsTextBox.Text = _currentMetrics;
                LastUpdatedText.Text = DateTime.Now.ToString("HH:mm:ss");
            }
            finally
            {
                RefreshButton.IsEnabled = true;
            }
        }
        
        private void CopyButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(_currentMetrics))
            {
                Clipboard.SetText(_currentMetrics);
                
                // Show temporary confirmation
                var originalContent = CopyButton.Content;
                CopyButton.Content = "Copied!";
                CopyButton.IsEnabled = false;
                
                var timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                timer.Tick += (s, args) =>
                {
                    CopyButton.Content = originalContent;
                    CopyButton.IsEnabled = true;
                    timer.Stop();
                };
                timer.Start();
            }
        }
        
        private void AutoRefreshCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            _refreshTimer.Start();
        }
        
        private void AutoRefreshCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            _refreshTimer.Stop();
        }
        
        protected override void OnClosed(EventArgs e)
        {
            _refreshTimer.Stop();
            base.OnClosed(e);
        }
    }
} 