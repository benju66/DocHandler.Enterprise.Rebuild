using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows;
using Serilog;
using ComIDataObject = System.Runtime.InteropServices.ComTypes.IDataObject;
using WpfIDataObject = System.Windows.IDataObject;

namespace DocHandler.Helpers
{
    /// <summary>
    /// Helper class for extracting attachments dragged from Outlook emails
    /// </summary>
    public static class OutlookAttachmentHelper
    {
        private static readonly ILogger _logger = Log.ForContext(typeof(OutlookAttachmentHelper));

        // Outlook data format names
        private const string CFSTR_FILEDESCRIPTORW = "FileGroupDescriptorW";
        private const string CFSTR_FILEDESCRIPTORA = "FileGroupDescriptor";
        private const string CFSTR_FILECONTENTS = "FileContents";

        /// <summary>
        /// Extracts Outlook attachments from drag data and saves them to temporary files
        /// </summary>
        public static List<string> ExtractOutlookAttachments(WpfIDataObject dataObject)
        {
            var extractedFiles = new List<string>();

            try
            {
                // Check for Unicode version first
                if (dataObject.GetDataPresent(CFSTR_FILEDESCRIPTORW))
                {
                    _logger.Information("Extracting Outlook attachments (Unicode)");
                    extractedFiles = ExtractAttachmentsUnicode(dataObject);
                }
                else if (dataObject.GetDataPresent(CFSTR_FILEDESCRIPTORA))
                {
                    _logger.Information("Extracting Outlook attachments (ANSI)");
                    extractedFiles = ExtractAttachmentsAnsi(dataObject);
                }
                else
                {
                    _logger.Debug("No Outlook attachment formats found in drag data");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract Outlook attachments");
            }

            return extractedFiles;
        }

        private static List<string> ExtractAttachmentsUnicode(WpfIDataObject dataObject)
        {
            var files = new List<string>();
            
            try
            {
                // Get the file group descriptor
                var descriptorData = dataObject.GetData(CFSTR_FILEDESCRIPTORW);
                if (descriptorData == null) return files;

                var descriptorStream = descriptorData as MemoryStream;
                if (descriptorStream == null)
                {
                    _logger.Warning("FileGroupDescriptorW data is not a MemoryStream");
                    return files;
                }

                var descriptorBytes = descriptorStream.ToArray();
                
                // First 4 bytes contain the count of files
                int fileCount = BitConverter.ToInt32(descriptorBytes, 0);
                _logger.Debug("Found {Count} attachments to extract", fileCount);

                if (fileCount == 0) return files;

                // Each FILEDESCRIPTORW is 592 bytes (with padding)
                int descriptorSize = Marshal.SizeOf(typeof(FILEDESCRIPTORW));

                for (int i = 0; i < fileCount; i++)
                {
                    try
                    {
                        // Calculate offset for this file descriptor
                        int offset = 4 + (i * descriptorSize);
                        
                        // Marshal the bytes to FILEDESCRIPTORW structure
                        IntPtr ptr = Marshal.AllocHGlobal(descriptorSize);
                        try
                        {
                            Marshal.Copy(descriptorBytes, offset, ptr, descriptorSize);
                            var descriptor = (FILEDESCRIPTORW)Marshal.PtrToStructure(ptr, typeof(FILEDESCRIPTORW))!;
                            
                            // Get the filename
                            string fileName = descriptor.cFileName;
                            if (string.IsNullOrEmpty(fileName)) continue;

                            _logger.Debug("Extracting attachment: {FileName}", fileName);

                            // Get the file contents
                            var fileData = GetFileContents(dataObject, i);
                            if (fileData != null && fileData.Length > 0)
                            {
                                // Save to temp file
                                string tempPath = SaveToTempFile(fileName, fileData);
                                if (!string.IsNullOrEmpty(tempPath))
                                {
                                    files.Add(tempPath);
                                    _logger.Information("Extracted attachment to: {Path}", tempPath);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(ptr);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to extract attachment at index {Index}", i);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to process Unicode file descriptors");
            }

            return files;
        }

        private static List<string> ExtractAttachmentsAnsi(WpfIDataObject dataObject)
        {
            var files = new List<string>();
            
            try
            {
                // Get the file group descriptor
                var descriptorData = dataObject.GetData(CFSTR_FILEDESCRIPTORA);
                if (descriptorData == null) return files;

                var descriptorStream = descriptorData as MemoryStream;
                if (descriptorStream == null)
                {
                    _logger.Warning("FileGroupDescriptor data is not a MemoryStream");
                    return files;
                }

                var descriptorBytes = descriptorStream.ToArray();
                
                // First 4 bytes contain the count of files
                int fileCount = BitConverter.ToInt32(descriptorBytes, 0);
                _logger.Debug("Found {Count} attachments to extract", fileCount);

                if (fileCount == 0) return files;

                // Each FILEDESCRIPTORA is 332 bytes (with padding)
                int descriptorSize = Marshal.SizeOf(typeof(FILEDESCRIPTORA));

                for (int i = 0; i < fileCount; i++)
                {
                    try
                    {
                        // Calculate offset for this file descriptor
                        int offset = 4 + (i * descriptorSize);
                        
                        // Marshal the bytes to FILEDESCRIPTORA structure
                        IntPtr ptr = Marshal.AllocHGlobal(descriptorSize);
                        try
                        {
                            Marshal.Copy(descriptorBytes, offset, ptr, descriptorSize);
                            var descriptor = (FILEDESCRIPTORA)Marshal.PtrToStructure(ptr, typeof(FILEDESCRIPTORA))!;
                            
                            // Get the filename (convert from ANSI)
                            string fileName = GetAnsiString(descriptor.cFileName);
                            if (string.IsNullOrEmpty(fileName)) continue;

                            _logger.Debug("Extracting attachment: {FileName}", fileName);

                            // Get the file contents
                            var fileData = GetFileContents(dataObject, i);
                            if (fileData != null && fileData.Length > 0)
                            {
                                // Save to temp file
                                string tempPath = SaveToTempFile(fileName, fileData);
                                if (!string.IsNullOrEmpty(tempPath))
                                {
                                    files.Add(tempPath);
                                    _logger.Information("Extracted attachment to: {Path}", tempPath);
                                }
                            }
                        }
                        finally
                        {
                            Marshal.FreeHGlobal(ptr);
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to extract attachment at index {Index}", i);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to process ANSI file descriptors");
            }

            return files;
        }

        private static byte[]? GetFileContents(WpfIDataObject dataObject, int index)
        {
            try
            {
                // Get the underlying COM IDataObject using reflection
                var comDataObject = GetComDataObject(dataObject);
                if (comDataObject != null)
                {
                    return GetFileContentsFromCom(comDataObject, index);
                }
                
                // Fallback: Try direct approach (works for some cases)
                var data = dataObject.GetData(CFSTR_FILECONTENTS);
                if (data is MemoryStream ms)
                {
                    // This might only work for single attachments
                    if (index == 0)
                    {
                        return ms.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to get file contents for index {Index}", index);
            }

            return null;
        }

        private static ComIDataObject? GetComDataObject(WpfIDataObject dataObject)
        {
            try
            {
                // WPF DataObject wraps the COM IDataObject
                // We need to use reflection to get the underlying COM object
                var dataObjectType = dataObject.GetType();
                var innerDataField = dataObjectType.GetField("_innerData", BindingFlags.NonPublic | BindingFlags.Instance);
                
                if (innerDataField != null)
                {
                    var innerData = innerDataField.GetValue(dataObject);
                    return innerData as ComIDataObject;
                }
                
                // Alternative: try direct cast (might work in some scenarios)
                return dataObject as ComIDataObject;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to get COM IDataObject from WPF DataObject");
                return null;
            }
        }

        private static byte[]? GetFileContentsFromCom(ComIDataObject comDataObject, int index)
        {
            try
            {
                var format = DataFormats.GetDataFormat(CFSTR_FILECONTENTS);
                
                var formatetc = new FORMATETC
                {
                    cfFormat = (short)format.Id,
                    dwAspect = DVASPECT.DVASPECT_CONTENT,
                    lindex = index,
                    tymed = TYMED.TYMED_ISTREAM | TYMED.TYMED_HGLOBAL
                };

                comDataObject.GetData(ref formatetc, out var medium);
                
                try
                {
                    if (medium.tymed == TYMED.TYMED_ISTREAM && medium.unionmember != IntPtr.Zero)
                    {
                        var stream = (IStream)Marshal.GetObjectForIUnknown(medium.unionmember);
                        return ReadStreamToBytes(stream);
                    }
                    else if (medium.tymed == TYMED.TYMED_HGLOBAL && medium.unionmember != IntPtr.Zero)
                    {
                        return ReadHGlobalToBytes(medium.unionmember);
                    }
                }
                finally
                {
                    ReleaseStgMedium(ref medium);
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to get file contents from COM for index {Index}", index);
            }

            return null;
        }

        private static byte[] ReadStreamToBytes(IStream stream)
        {
            var bytes = new List<byte>();
            var buffer = new byte[4096];
            
            while (true)
            {
                stream.Read(buffer, buffer.Length, IntPtr.Zero);
                stream.Stat(out var stat, 0);
                
                var bytesRead = (int)(stat.cbSize % buffer.Length);
                if (bytesRead == 0 && stat.cbSize > 0)
                    bytesRead = buffer.Length;
                
                if (bytesRead == 0) break;
                
                for (int i = 0; i < bytesRead; i++)
                {
                    bytes.Add(buffer[i]);
                }
                
                if (bytes.Count >= stat.cbSize)
                    break;
            }
            
            return bytes.ToArray();
        }

        private static byte[] ReadHGlobalToBytes(IntPtr hGlobal)
        {
            var size = GlobalSize(hGlobal).ToInt32();
            var bytes = new byte[size];
            var ptr = GlobalLock(hGlobal);
            
            try
            {
                Marshal.Copy(ptr, bytes, 0, size);
            }
            finally
            {
                GlobalUnlock(hGlobal);
            }
            
            return bytes;
        }

        private static string GetAnsiString(byte[] bytes)
        {
            // Find null terminator
            int nullIndex = Array.IndexOf<byte>(bytes, 0);
            if (nullIndex == -1) nullIndex = bytes.Length;
            
            return Encoding.Default.GetString(bytes, 0, nullIndex);
        }

        private static string SaveToTempFile(string originalFileName, byte[] fileData)
        {
            try
            {
                // Create temp directory for Outlook attachments
                var tempDir = Path.Combine(Path.GetTempPath(), "DocHandler", "OutlookAttachments");
                Directory.CreateDirectory(tempDir);

                // Generate unique filename
                var fileName = Path.GetFileName(originalFileName);
                var uniqueFileName = $"{Path.GetFileNameWithoutExtension(fileName)}_{Guid.NewGuid():N}{Path.GetExtension(fileName)}";
                var tempPath = Path.Combine(tempDir, uniqueFileName);

                // Write file
                File.WriteAllBytes(tempPath, fileData);
                
                return tempPath;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save temp file for: {FileName}", originalFileName);
                return string.Empty;
            }
        }

        /// <summary>
        /// Cleans up temporary files created during Outlook attachment extraction
        /// </summary>
        public static void CleanupTempFiles()
        {
            try
            {
                var tempDir = Path.Combine(Path.GetTempPath(), "DocHandler", "OutlookAttachments");
                if (Directory.Exists(tempDir))
                {
                    // Delete files older than 24 hours
                    var cutoffTime = DateTime.Now.AddHours(-24);
                    var files = Directory.GetFiles(tempDir);
                    
                    foreach (var file in files)
                    {
                        try
                        {
                            var fileInfo = new FileInfo(file);
                            if (fileInfo.CreationTime < cutoffTime)
                            {
                                File.Delete(file);
                                _logger.Debug("Cleaned up old temp file: {File}", file);
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to cleanup temp files");
            }
        }

        #region Native Methods
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GlobalUnlock(IntPtr hMem);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GlobalSize(IntPtr hMem);

        [DllImport("ole32.dll")]
        private static extern void ReleaseStgMedium(ref STGMEDIUM pmedium);
        #endregion

        #region File Descriptor Structures
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct FILEDESCRIPTORW
        {
            public uint dwFlags;
            public Guid clsid;
            public SIZEL sizel;
            public POINTL pointl;
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string cFileName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
        private struct FILEDESCRIPTORA
        {
            public uint dwFlags;
            public Guid clsid;
            public SIZEL sizel;
            public POINTL pointl;
            public uint dwFileAttributes;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftCreationTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastAccessTime;
            public System.Runtime.InteropServices.ComTypes.FILETIME ftLastWriteTime;
            public uint nFileSizeHigh;
            public uint nFileSizeLow;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 260)]
            public byte[] cFileName;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SIZEL
        {
            public int cx;
            public int cy;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct POINTL
        {
            public int x;
            public int y;
        }
        #endregion
    }
}