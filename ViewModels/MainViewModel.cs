using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using DocHandler.Services;
using Serilog;

namespace DocHandler.ViewModels
{
    public partial class MainViewModel : ObservableObject
    {
        private readonly ILogger _logger;
        private readonly FileProcessingService _fileProcessingService;
        private readonly ConfigurationService _configService;
        private readonly OfficeConversionService _officeConversionService;
        
        public ConfigurationService ConfigService => _configService;
        
        [ObservableProperty]
        private ObservableCollection<FileItem> _pendingFiles = new();
        
        [ObservableProperty]
        private bool _isProcessing;
        
        [ObservableProperty]
        private double _progressValue;
        
        [ObservableProperty]
        private string _statusMessage = "Drop files here to begin";
        
        [ObservableProperty]
        private bool _canProcess;
        
        [ObservableProperty]
        private string _processButtonText = "Process Files";
        
        public MainViewModel()
        {
            _logger = Log.ForContext<MainViewModel>();
            _fileProcessingService = new FileProcessingService();
            _configService = new ConfigurationService();
            _officeConversionService = new OfficeConversionService();
            
            // Update UI when files are added/removed
            PendingFiles.CollectionChanged += (s, e) => UpdateUI();
        }
        
        private void UpdateUI()
        {
            CanProcess = PendingFiles.Count > 0 && !IsProcessing;
            ProcessButtonText = PendingFiles.Count > 1 ? "Merge and Save" : "Process Files";
            
            if (PendingFiles.Count == 0)
            {
                StatusMessage = "Drop files here to begin";
            }
            else
            {
                StatusMessage = $"{PendingFiles.Count} file(s) ready to process";
            }
        }
        
        public void AddFiles(string[] filePaths)
        {
            var validFiles = _fileProcessingService.ValidateDroppedFiles(filePaths);
            
            foreach (var file in validFiles)
            {
                // Check if file already added
                if (PendingFiles.Any(f => f.FilePath == file))
                {
                    _logger.Information("File already in list: {FilePath}", file);
                    continue;
                }
                
                var fileItem = new FileItem
                {
                    FilePath = file,
                    FileName = Path.GetFileName(file),
                    FileSize = new FileInfo(file).Length,
                    FileType = Path.GetExtension(file).ToUpperInvariant().TrimStart('.')
                };
                
                PendingFiles.Add(fileItem);
            }
            
            if (validFiles.Count != filePaths.Length)
            {
                var invalidCount = filePaths.Length - validFiles.Count;
                MessageBox.Show($"{invalidCount} file(s) were not added because they are not supported.", 
                    "Some Files Not Added", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        
        [RelayCommand]
        private async Task ProcessFiles()
        {
            if (PendingFiles.Count == 0) return;
            
            try
            {
                IsProcessing = true;
                ProgressValue = 0;
                StatusMessage = "Processing files...";
                
                // Get output folder
                var outputPath = _fileProcessingService.CreateOutputFolder(_configService.Config.DefaultSaveLocation);
                
                // Process each file
                var totalFiles = PendingFiles.Count;
                var processedCount = 0;
                
                foreach (var fileItem in PendingFiles.ToList())
                {
                    StatusMessage = $"Processing {fileItem.FileName}...";
                    
                    // Check if conversion is needed
                    var extension = Path.GetExtension(fileItem.FilePath).ToLowerInvariant();
                    var outputFileName = Path.GetFileNameWithoutExtension(fileItem.FilePath) + ".pdf";
                    var outputFilePath = Path.Combine(outputPath, outputFileName);
                    
                    if (extension == ".doc" || extension == ".docx")
                    {
                        // Convert Word to PDF
                        var result = await _officeConversionService.ConvertWordToPdf(fileItem.FilePath, outputFilePath);
                        if (!result.Success)
                        {
                            _logger.Error("Failed to convert {File}: {Error}", fileItem.FileName, result.ErrorMessage);
                            MessageBox.Show($"Failed to convert {fileItem.FileName}: {result.ErrorMessage}", 
                                "Conversion Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                        }
                    }
                    else
                    {
                        // Just copy PDFs and other files for now
                        await Task.Run(() => 
                        {
                            var destPath = Path.Combine(outputPath, fileItem.FileName);
                            File.Copy(fileItem.FilePath, destPath, overwrite: false);
                        });
                    }
                    
                    processedCount++;
                    ProgressValue = (double)processedCount / totalFiles * 100;
                }
                
                // Update configuration with recent location
                _configService.AddRecentLocation(outputPath);
                
                StatusMessage = $"Completed! Files saved to: {outputPath}";
                
                // Open output folder
                System.Diagnostics.Process.Start("explorer.exe", outputPath);
                
                // Clear the list after successful processing
                PendingFiles.Clear();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error processing files");
                MessageBox.Show($"An error occurred: {ex.Message}", "Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                StatusMessage = "Error occurred during processing";
            }
            finally
            {
                IsProcessing = false;
                ProgressValue = 0;
                UpdateUI();
            }
        }
        
        [RelayCommand]
        private void ClearFiles()
        {
            PendingFiles.Clear();
            StatusMessage = "Files cleared";
            UpdateUI();
        }
        
        [RelayCommand]
        private void RemoveFile(FileItem? fileItem)
        {
            if (fileItem != null)
            {
                PendingFiles.Remove(fileItem);
                UpdateUI();
            }
        }
        
        public void Cleanup()
        {
            _officeConversionService?.Dispose();
        }
        
        public void SaveWindowState(double left, double top, double width, double height, string state)
        {
            if (_configService.Config.RememberWindowPosition)
            {
                _configService.UpdateWindowPosition(left, top, width, height, state);
                _ = _configService.SaveConfiguration();
            }
        }
    }
    
    public class FileItem
    {
        public string FilePath { get; set; } = "";
        public string FileName { get; set; } = "";
        public string FileType { get; set; } = "";
        public long FileSize { get; set; }
        
        public string FileSizeDisplay => FormatFileSize(FileSize);
        
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