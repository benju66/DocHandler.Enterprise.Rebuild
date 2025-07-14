using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;

namespace DocHandler.ViewModels
{
    public partial class QueueDetailsViewModel : ObservableObject
    {
        private readonly SaveQuotesQueueService _queueService;
        
        public ObservableCollection<SaveQuoteItem> QueueItems => _queueService.AllItems;
        
        public QueueDetailsViewModel(SaveQuotesQueueService queueService)
        {
            _queueService = queueService;
        }
        
        [RelayCommand]
        private void RemoveItem(SaveQuoteItem item)
        {
            if (item?.Status == SaveQuoteStatus.Queued)
            {
                _queueService.CancelItem(item);
            }
        }
        
        [RelayCommand]
        private void ClearCompleted()
        {
            _queueService.ClearCompleted();
        }
    }
} 