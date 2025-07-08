using CommunityToolkit.Mvvm.ComponentModel;

namespace DocHandler.Models
{
    public enum ValidationStatus
    {
        Pending,
        Validating,
        Valid,
        Invalid
    }
    
    public partial class FileItem : ObservableObject
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public long FileSize { get; set; }
        public string FileType { get; set; } = "";
        
        [ObservableProperty]
        private ValidationStatus _validationStatus = ValidationStatus.Pending;
        
        [ObservableProperty]
        private string? _validationError;
        
        [ObservableProperty]
        private bool _isProcessing;
        
        public string DisplayFileSize => FormatFileSize(FileSize);
        
        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }
} 