using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using DocHandler.Services;

namespace DocHandler.Helpers
{
    public class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility visibility)
            {
                return visibility == Visibility.Visible;
            }
            return false;
        }
    }
    
    public class IntToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intValue)
            {
                int compareValue = 0;
                if (parameter != null && int.TryParse(parameter.ToString(), out int paramValue))
                {
                    compareValue = paramValue;
                }
                
                return intValue == compareValue ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    
    public class IntToBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int intValue)
            {
                return intValue > 0;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    
    public class AutoScanMenuHeaderConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isEnabled)
            {
                return isEnabled ? "✓ Auto-Scan Company Names" : "Auto-Scan Company Names";
            }
            return "Auto-Scan Company Names";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    
    public class MultiBooleanConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null) return false;
            
            foreach (var value in values)
            {
                if (value is bool boolValue && !boolValue)
                    return false;
            }
            
            return true;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    
    /// <summary>
    /// Formats file paths for display with intelligent truncation and ">" separators
    /// Uses ">" symbols for better visual appeal and preserves the last 4-5 folders
    /// Prioritizes showing the complete final folder name without any truncation
    /// Example: "C:\Users\Name\OneDrive\Documents\Projects\2025\Project\Budgets\Quotes\Set" 
    /// becomes: "...> Projects > 2025 > Project > Budgets > Quotes > Set"
    /// </summary>
    public class PathDisplayConverter : IValueConverter
    {
        private const int MaxDisplayLength = 50; // Conservative character limit to ensure final folder shows completely
        private const int MinFoldersToShow = 3;  // Reduced minimum to prioritize final folder visibility
        private const int MaxFoldersToShow = 5;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is not string path || string.IsNullOrWhiteSpace(path))
                return string.Empty;

            try
            {
                // Split the path into parts
                var parts = path.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                               .Where(p => !string.IsNullOrWhiteSpace(p))
                               .ToArray();

                if (parts.Length == 0)
                    return string.Empty;

                // Create display version with ">" separators
                var fullDisplayPath = string.Join(" > ", parts);

                // If the full path fits comfortably, show it all
                if (fullDisplayPath.Length <= MaxDisplayLength)
                {
                    return fullDisplayPath;
                }

                // Path is too long, apply intelligent truncation
                // Start with just the final folder to ensure it's always fully visible
                var finalFolder = parts[parts.Length - 1];
                var baseTruncatedPath = "...> " + finalFolder;
                
                // If even the final folder with prefix is too long, just show the final folder
                if (baseTruncatedPath.Length > MaxDisplayLength)
                {
                    return finalFolder;
                }

                // Now try to add more folders working backwards, but never exceed our limit
                for (int folderCount = 2; folderCount <= Math.Min(MaxFoldersToShow, parts.Length); folderCount++)
                {
                    var lastParts = parts.Skip(parts.Length - folderCount).ToArray();
                    var candidatePath = (parts.Length > folderCount ? "...> " : "") + string.Join(" > ", lastParts);
                    
                    // If adding this folder would make it too long, use the previous version
                    if (candidatePath.Length > MaxDisplayLength)
                    {
                        // Return the previous working version
                        var previousParts = parts.Skip(parts.Length - (folderCount - 1)).ToArray();
                        return "...> " + string.Join(" > ", previousParts);
                    }
                    
                    // If this is the last iteration and it fits, use it
                    if (folderCount == Math.Min(MaxFoldersToShow, parts.Length))
                    {
                        return candidatePath;
                    }
                }

                // Fallback: just show the final folder if everything else fails
                return finalFolder;
            }
            catch
            {
                // Fallback to original path if any error occurs
                return path;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    
    public class StatusToIconConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SaveQuoteStatus status)
            {
                return status switch
                {
                    SaveQuoteStatus.Queued => "⏳",
                    SaveQuoteStatus.Processing => "⚡",
                    SaveQuoteStatus.Completed => "✓",
                    SaveQuoteStatus.Failed => "❌",
                    SaveQuoteStatus.Cancelled => "⛔",
                    _ => "?"
                };
            }
            return "?";
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class QueuedStatusToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is SaveQuoteStatus status)
            {
                return status == SaveQuoteStatus.Queued ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class QueueStatusConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values.Length != 4) return "";
            
            var isProcessing = values[0] as bool? ?? false;
            var processedCount = values[1] as int? ?? 0;
            var totalCount = values[2] as int? ?? 0;
            var completionMessage = values[3] as string ?? "";
            
            if (isProcessing)
            {
                return $"Saving {processedCount} of {totalCount}...";
            }
            else if (!string.IsNullOrEmpty(completionMessage))
            {
                return completionMessage;
            }
            
            return "";
        }
        
        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}