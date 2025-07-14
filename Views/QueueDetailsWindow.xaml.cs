using System;
using System.ComponentModel;
using System.Windows;
using DocHandler.Services;
using DocHandler.ViewModels;

namespace DocHandler.Views
{
    public partial class QueueDetailsWindow : Window
    {
        private readonly SaveQuotesQueueService _queueService;
        private readonly ConfigurationService _configService;
        private bool _isClosing = false;
        
        public QueueDetailsWindow(SaveQuotesQueueService queueService, ConfigurationService configService)
        {
            InitializeComponent();
            _queueService = queueService;
            _configService = configService;
            DataContext = new QueueDetailsViewModel(queueService);
            
            // Restore window size if saved
            if (_configService.Config.QueueWindowWidth.HasValue)
            {
                Width = _configService.Config.QueueWindowWidth.Value;
                Height = _configService.Config.QueueWindowHeight.Value;
            }
        }
        
        protected override void OnLocationChanged(EventArgs e)
        {
            base.OnLocationChanged(e);
            
            if (!_isClosing && WindowState == WindowState.Normal)
            {
                SaveWindowState();
            }
        }
        
        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);
            
            if (!_isClosing && WindowState == WindowState.Normal)
            {
                SaveWindowState();
            }
        }
        
        private void SaveWindowState()
        {
            _configService.Config.QueueWindowLeft = Left;
            _configService.Config.QueueWindowTop = Top;
            _configService.Config.QueueWindowWidth = Width;
            _configService.Config.QueueWindowHeight = Height;
            _configService.Config.QueueWindowIsOpen = true;
        }
        
        protected override void OnClosing(CancelEventArgs e)
        {
            _isClosing = true;
            base.OnClosing(e);
        }
        
        protected override void OnClosed(EventArgs e)
        {
            // Mark window as closed
            _configService.Config.QueueWindowIsOpen = false;
            _ = _configService.SaveConfiguration();
            
            base.OnClosed(e);
        }
        
        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
} 