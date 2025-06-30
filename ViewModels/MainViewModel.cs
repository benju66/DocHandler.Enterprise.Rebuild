using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
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

        // Checkbox options - hardcoded for now
        private bool _convertOfficeToPdf = true;
        public bool ConvertOfficeToPdf 
        { 
            get => _convertOfficeToPdf;
            set => SetProperty(ref _convertOfficeToPdf, value);
        }
        
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
            if (!PendingFiles.Any())
            {
                StatusMessage = "No files selected";
                return;
            }

            IsProcessing = true;
            StatusMessage = PendingFiles.Count > 1 ? "Merging and processing files..." : "Processing file...";

            try
            {
                var filePaths = PendingFiles.Select(f => f.FilePath).ToList();
                var outputDir = _configService.Config.DefaultSaveLocation;

                // Create output folder with timestamp
                outputDir = _fileProcessingService.CreateOutputFolder(outputDir);

                var result = await _fileProcessingService.ProcessFiles(filePaths, outputDir, ConvertOfficeToPdf);

                if (result.Success)
                {
                    if (result.IsMerged)
                    {
                        StatusMessage = $"Successfully merged {filePaths.Count} files into {Path.GetFileName(result.SuccessfulFiles.First())}";
                        _logger.Information("Files merged successfully");
                    }
                    else
                    {
                        StatusMessage = $"Successfully processed {result.SuccessfulFiles.Count} file(s)";
                        _logger.Information("Files processed successfully");
                    }

                    // Clear the file list after successful processing
                    PendingFiles.Clear();

                    // Update configuration with recent location
                    _configService.AddRecentLocation(outputDir);

                    // Open the output folder
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = outputDir,
                            UseShellExecute = true,
                            Verb = "open"
                        });
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to open output folder");
                    }
                }
                else
                {
                    var errorMessage = !string.IsNullOrEmpty(result.ErrorMessage) 
                        ? result.ErrorMessage 
                        : "Processing failed";
                    
                    StatusMessage = $"Error: {errorMessage}";
                    _logger.Error("File processing failed: {Error}", errorMessage);

                    if (result.FailedFiles.Any())
                    {
                        var failedFilesList = string.Join("\n", result.FailedFiles.Select(f => 
                            $"• {Path.GetFileName(f.FilePath)}: {f.Error}"));
                        
                        MessageBox.Show(
                            $"The following files could not be processed:\n\n{failedFilesList}",
                            "Processing Errors",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                StatusMessage = $"Error: {ex.Message}";
                _logger.Error(ex, "Unexpected error during file processing");
                MessageBox.Show(
                    $"An unexpected error occurred:\n\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
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