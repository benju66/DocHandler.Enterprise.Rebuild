using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Manages COM object lifecycle with automatic cleanup.
    /// Following SRP - single responsibility is COM resource management.
    /// </summary>
    public class ComResourceScope : IDisposable
    {
        private readonly List<(object comObject, string type, string context)> _trackedObjects = new();
        private readonly ILogger _logger = Log.ForContext<ComResourceScope>();
        private bool _disposed;

        /// <summary>
        /// Tracks a COM object for automatic cleanup
        /// </summary>
        /// <typeparam name="T">Type of COM object</typeparam>
        /// <param name="comObject">The COM object to track</param>
        /// <param name="objectType">Type description for logging</param>
        /// <param name="context">Context information for debugging</param>
        /// <returns>The same COM object for chaining</returns>
        public T Track<T>(T comObject, string objectType, string context = "") where T : class
        {
            if (comObject != null)
            {
                _trackedObjects.Add((comObject, objectType, context));
                ComHelper.TrackComObjectCreation(objectType, context);
                _logger.Debug("Tracked COM object {ObjectType} in {Context}", objectType, context);
            }
            return comObject;
        }

        /// <summary>
        /// Opens a Word document with automatic Documents collection cleanup
        /// </summary>
        /// <param name="wordApp">Word application instance</param>
        /// <param name="filePath">Path to the document</param>
        /// <param name="readOnly">Whether to open as read-only</param>
        /// <returns>The opened document</returns>
        public dynamic OpenWordDocument(dynamic wordApp, string filePath, bool readOnly = true)
        {
            var documents = Track(wordApp.Documents, "Documents", "OpenDocument");
            
            try
            {
                // Try modern parameters first (Word 2010+)
                try
                {
                    return Track(documents.Open(
                        filePath,
                        ReadOnly: readOnly,
                        AddToRecentFiles: false,
                        Repair: false,
                        ShowRepairs: false,
                        OpenAndRepair: false,
                        NoEncodingDialog: true,
                        Revert: false
                    ), "Document", "OpenDocument");
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                {
                    _logger.Debug("Extended Open parameters not supported, using basic Open");
                    // Fallback to basic Open for older Word versions
                    return Track(documents.Open(filePath, ReadOnly: readOnly), "Document", "OpenDocument");
                }
            }
            catch
            {
                // Re-throw but Documents will still be cleaned up
                throw;
            }
        }

        /// <summary>
        /// Opens an Excel workbook with automatic Workbooks collection cleanup
        /// </summary>
        /// <param name="excelApp">Excel application instance</param>
        /// <param name="filePath">Path to the workbook</param>
        /// <param name="readOnly">Whether to open as read-only</param>
        /// <returns>The opened workbook</returns>
        public dynamic OpenExcelWorkbook(dynamic excelApp, string filePath, bool readOnly = true)
        {
            var workbooks = Track(excelApp.Workbooks, "Workbooks", "OpenWorkbook");
            return Track(workbooks.Open(filePath, ReadOnly: readOnly), "Workbook", "OpenWorkbook");
        }

        /// <summary>
        /// Gets Documents collection with tracking for cleanup
        /// </summary>
        /// <param name="wordApp">Word application instance</param>
        /// <param name="context">Context for debugging</param>
        /// <returns>Tracked Documents collection</returns>
        public dynamic GetDocuments(dynamic wordApp, string context = "")
        {
            return Track(wordApp.Documents, "Documents", context);
        }

        /// <summary>
        /// Gets Workbooks collection with tracking for cleanup
        /// </summary>
        /// <param name="excelApp">Excel application instance</param>
        /// <param name="context">Context for debugging</param>
        /// <returns>Tracked Workbooks collection</returns>
        public dynamic GetWorkbooks(dynamic excelApp, string context = "")
        {
            return Track(excelApp.Workbooks, "Workbooks", context);
        }

        /// <summary>
        /// Releases all tracked COM objects in reverse order
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;

            _logger.Debug("Disposing ComResourceScope with {Count} tracked objects", _trackedObjects.Count);

            // Release all tracked objects in reverse order (LIFO)
            for (int i = _trackedObjects.Count - 1; i >= 0; i--)
            {
                var (comObject, type, context) = _trackedObjects[i];
                try
                {
                    ComHelper.SafeReleaseComObject(comObject, type, context);
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error releasing COM object {Type} in {Context}", type, context);
                }
            }

            _trackedObjects.Clear();
            _disposed = true;
            _logger.Debug("ComResourceScope disposed successfully");
        }
    }
} 