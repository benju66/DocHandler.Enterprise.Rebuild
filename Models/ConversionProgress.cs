using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace DocHandler.Models
{
    /// <summary>
    /// Tracks detailed progress of document conversion operations.
    /// </summary>
    public class ConversionProgress : INotifyPropertyChanged
    {
        private int _totalFiles;
        private int _completedFiles;
        private int _failedFiles;
        private int _activeConversions;
        private string _currentFileName = string.Empty;
        private string _currentOperation = "Ready";
        private double _fileProgress;
        private DateTime _startTime;
        private DateTime? _lastUpdateTime;

        public int TotalFiles
        {
            get => _totalFiles;
            set
            {
                if (_totalFiles != value)
                {
                    _totalFiles = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(OverallProgress));
                    OnPropertyChanged(nameof(RemainingFiles));
                    OnPropertyChanged(nameof(DetailedStatus));
                    OnPropertyChanged(nameof(EstimatedTimeRemaining));
                }
            }
        }

        public int CompletedFiles
        {
            get => _completedFiles;
            set
            {
                if (_completedFiles != value)
                {
                    _completedFiles = value;
                    _lastUpdateTime = DateTime.Now;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(OverallProgress));
                    OnPropertyChanged(nameof(RemainingFiles));
                    OnPropertyChanged(nameof(DetailedStatus));
                    OnPropertyChanged(nameof(EstimatedTimeRemaining));
                }
            }
        }

        public int FailedFiles
        {
            get => _failedFiles;
            set
            {
                if (_failedFiles != value)
                {
                    _failedFiles = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DetailedStatus));
                }
            }
        }

        public int ActiveConversions
        {
            get => _activeConversions;
            set
            {
                if (_activeConversions != value)
                {
                    _activeConversions = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DetailedStatus));
                }
            }
        }

        public string CurrentFileName
        {
            get => _currentFileName;
            set
            {
                if (_currentFileName != value)
                {
                    _currentFileName = value ?? string.Empty;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(DetailedStatus));
                }
            }
        }

        public string CurrentOperation
        {
            get => _currentOperation;
            set
            {
                if (_currentOperation != value)
                {
                    _currentOperation = value ?? "Ready";
                    OnPropertyChanged();
                }
            }
        }

        public double FileProgress
        {
            get => _fileProgress;
            set
            {
                if (Math.Abs(_fileProgress - value) > 0.01)
                {
                    _fileProgress = value;
                    OnPropertyChanged();
                }
            }
        }

        public int RemainingFiles => Math.Max(0, TotalFiles - CompletedFiles - FailedFiles);

        public double OverallProgress => TotalFiles > 0 
            ? Math.Min(100, (double)(CompletedFiles + FailedFiles) / TotalFiles * 100) 
            : 0;

        public string DetailedStatus
        {
            get
            {
                if (TotalFiles == 0)
                    return "Ready";

                if (CompletedFiles + FailedFiles >= TotalFiles)
                    return $"Completed: {CompletedFiles} successful, {FailedFiles} failed";

                if (ActiveConversions > 1)
                    return $"Processing {ActiveConversions} files ({CompletedFiles}/{TotalFiles})";

                if (!string.IsNullOrEmpty(CurrentFileName))
                    return $"{CurrentFileName} ({CompletedFiles + 1}/{TotalFiles})";

                return $"Processing {CompletedFiles}/{TotalFiles}";
            }
        }

        public string EstimatedTimeRemaining
        {
            get
            {
                if (CompletedFiles == 0 || !_lastUpdateTime.HasValue)
                    return string.Empty;

                var elapsed = DateTime.Now - _startTime;
                var avgTimePerFile = elapsed.TotalSeconds / CompletedFiles;
                var remainingSeconds = avgTimePerFile * RemainingFiles;

                if (remainingSeconds < 60)
                    return $"{(int)remainingSeconds}s remaining";
                else if (remainingSeconds < 3600)
                    return $"{(int)(remainingSeconds / 60)}m remaining";
                else
                    return $"{remainingSeconds / 3600:F1}h remaining";
            }
        }

        public void StartBatch(int totalFiles)
        {
            TotalFiles = totalFiles;
            CompletedFiles = 0;
            FailedFiles = 0;
            ActiveConversions = 0;
            CurrentFileName = string.Empty;
            CurrentOperation = "Initializing...";
            FileProgress = 0;
            _startTime = DateTime.Now;
            _lastUpdateTime = null;
        }

        public void UpdateFileProgress(string fileName, string operation, double progress = 0)
        {
            CurrentFileName = fileName;
            CurrentOperation = operation;
            FileProgress = progress;
        }

        public void CompleteFile(bool success)
        {
            if (success)
                CompletedFiles++;
            else
                FailedFiles++;

            FileProgress = 0;
        }

        public void Reset()
        {
            TotalFiles = 0;
            CompletedFiles = 0;
            FailedFiles = 0;
            ActiveConversions = 0;
            CurrentFileName = string.Empty;
            CurrentOperation = "Ready";
            FileProgress = 0;
            _startTime = DateTime.Now;
            _lastUpdateTime = null;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
} 