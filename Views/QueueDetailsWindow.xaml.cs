using System;
using System.Windows;
using DocHandler.Services;
using DocHandler.ViewModels;

namespace DocHandler.Views
{
    public partial class QueueDetailsWindow : Window
    {
        private readonly SaveQuotesQueueService _queueService;
        
        public QueueDetailsWindow(SaveQuotesQueueService queueService)
        {
            InitializeComponent();
            _queueService = queueService;
            DataContext = new QueueDetailsViewModel(queueService);
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        
        protected override void OnClosed(EventArgs e)
        {
            // Save window position if needed
            if (Owner is MainWindow mainWindow && mainWindow.DataContext is MainViewModel mainViewModel)
            {
                var configService = mainViewModel.ConfigService;
                configService.Config.QueueWindowLeft = Left;
                configService.Config.QueueWindowTop = Top;
                _ = configService.SaveConfiguration();
            }
            
            base.OnClosed(e);
        }
    }
} 