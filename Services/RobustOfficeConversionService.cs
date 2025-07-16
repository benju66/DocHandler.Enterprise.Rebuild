// Folder: Services/
// File: RobustOfficeConversionService.cs
// Robust Office conversion service with proper COM threading and fallback mechanisms
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Serilog;

namespace DocHandler.Services
{
    public class RobustOfficeConversionService : IDisposable
    {
        private readonly ILogger _logger;
        private readonly SemaphoreSlim _conversionSemaphore;
        private readonly int _maxConcurrency;
        private readonly Timer _cleanupTimer;
        private bool _disposed;
        private bool? _officeAvailable;
        private int _conversionsCount;
        private readonly object _cleanupLock = new object();
        private bool _usePooling = true;

        // Simple Word app pool - create on demand, dispose when unhealthy
        private readonly ConcurrentQueue<WordAppInfo> _wordAppPool = new();

        public RobustOfficeConversionService(int maxConcurrency = 0)
        {
            _logger = Log.ForContext<RobustOfficeConversionService>();
            
            // Conservative concurrency
            _maxConcurrency = maxConcurrency > 0 ? maxConcurrency : Math.Max(1, Environment.ProcessorCount - 2);
            _conversionSemaphore = new SemaphoreSlim(_maxConcurrency);
            
            _logger.Information("Initializing robust Office conversion service with max concurrency: {MaxConcurrency}", _maxConcurrency);
            
            // Cleanup timer - every 5 minutes
            _cleanupTimer = new Timer(PerformCleanup, null, 
                TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
        }

        private bool IsOfficeAvailable()
        {
            if (_officeAvailable.HasValue)
                return _officeAvailable.Value;
                
            try
            {
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType != null)
                {
                    dynamic testApp = null;
                    try
                    {
                        testApp = Activator.CreateInstance(wordType);
                        testApp.Visible = false;
                        testApp.Quit();
                        _officeAvailable = true;
                        return true;
                    }
                    finally
                    {
                        if (testApp != null)
                        {
                            try
                            {
                                ComHelper.SafeReleaseComObject(testApp, "WordApp", "RobustOfficeAvailabilityCheck");
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Microsoft Office is not available");
            }
            
            _officeAvailable = false;
            return false;
        }

        public async Task<ConversionResult> ConvertWordToPdf(string inputPath, string outputPath)
        {
            if (!IsOfficeAvailable())
            {
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = "Microsoft Office is not installed or accessible. Please install Microsoft Office to convert Word documents to PDF."
                };
            }

            await _conversionSemaphore.WaitAsync();
            
            try
            {
                var result = _usePooling ? 
                    await ConvertWithPooling(inputPath, outputPath) :
                    await ConvertWithNewInstance(inputPath, outputPath);
                
                Interlocked.Increment(ref _conversionsCount);
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Conversion failed for {File}", Path.GetFileName(inputPath));
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Conversion error: {ex.Message}"
                };
            }
            finally
            {
                _conversionSemaphore.Release();
            }
        }

        private async Task<ConversionResult> ConvertWithPooling(string inputPath, string outputPath)
        {
            WordAppInfo? wordInfo = null;
            
            try
            {
                // Try to get Word app from pool
                if (!_wordAppPool.TryDequeue(out wordInfo))
                {
                    wordInfo = await CreateWordAppInfo();
                    if (wordInfo == null)
                    {
                        _logger.Warning("Failed to create Word application, falling back to new instance mode");
                        _usePooling = false;
                        return await ConvertWithNewInstance(inputPath, outputPath);
                    }
                }

                // Perform conversion
                var result = await ConvertUsingWordApp(wordInfo, inputPath, outputPath);
                
                // Return healthy instance to pool, dispose unhealthy ones
                if (result.Success && wordInfo.IsHealthy)
                {
                    _wordAppPool.Enqueue(wordInfo);
                    wordInfo = null; // Don't dispose
                }
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Pooled conversion failed, falling back to new instance");
                _usePooling = false;
                return await ConvertWithNewInstance(inputPath, outputPath);
            }
            finally
            {
                // Dispose if not returned to pool
                if (wordInfo != null)
                {
                    await DisposeWordAppInfo(wordInfo);
                }
            }
        }

        private async Task<ConversionResult> ConvertWithNewInstance(string inputPath, string outputPath)
        {
            WordAppInfo? wordInfo = null;
            
            try
            {
                wordInfo = await CreateWordAppInfo();
                if (wordInfo == null)
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Could not create Word application instance"
                    };
                }

                return await ConvertUsingWordApp(wordInfo, inputPath, outputPath);
            }
            finally
            {
                if (wordInfo != null)
                {
                    await DisposeWordAppInfo(wordInfo);
                }
            }
        }

        private async Task<WordAppInfo?> CreateWordAppInfo()
        {
            return await Task.Run(() =>
            {
                try
                {
                    // Ensure STA thread for COM
                    if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    {
                        Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                    }

                    Type? wordType = Type.GetTypeFromProgID("Word.Application");
                    if (wordType == null)
                    {
                        _logger.Error("Word.Application ProgID not found");
                        return null;
                    }

                    dynamic wordApp = Activator.CreateInstance(wordType);
                    if (wordApp == null)
                    {
                        _logger.Error("Failed to create Word application instance");
                        return null;
                    }

                    // Basic optimizations
                    wordApp.Visible = false;
                    wordApp.DisplayAlerts = 0;
                    wordApp.ScreenUpdating = false;
                    
                    // Get process ID
                    var processId = (int)wordApp.GetType().InvokeMember("ProcessID", 
                        System.Reflection.BindingFlags.GetProperty, null, wordApp, null);

                    var wordInfo = new WordAppInfo
                    {
                        Application = wordApp,
                        ProcessId = processId,
                        ThreadId = Thread.CurrentThread.ManagedThreadId,
                        CreatedAt = DateTime.UtcNow,
                        IsHealthy = true
                    };

                    _logger.Debug("Created Word application with PID {ProcessId} on thread {ThreadId}", 
                        processId, wordInfo.ThreadId);
                    
                    return wordInfo;
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to create Word application");
                    return null;
                }
            });
        }

        private async Task<ConversionResult> ConvertUsingWordApp(WordAppInfo wordInfo, string inputPath, string outputPath)
        {
            return await Task.Run(() =>
            {
                var result = new ConversionResult();
                dynamic? doc = null;
                var stopwatch = Stopwatch.StartNew();
                
                try
                {
                    // Handle network paths
                    var workingInputPath = inputPath;
                    string? tempFilePath = null;
                    
                    if (IsNetworkPath(inputPath))
                    {
                        tempFilePath = Path.Combine(Path.GetTempPath(), $"DocHandler_{Guid.NewGuid()}.docx");
                        File.Copy(inputPath, tempFilePath, true);
                        workingInputPath = tempFilePath;
                        _logger.Debug("Using local copy for network path: {TempPath}", tempFilePath);
                    }

                    try
                    {
                        // Open document
                        doc = wordInfo.Application.Documents.Open(
                            workingInputPath,
                            ReadOnly: true,
                            AddToRecentFiles: false,
                            Repair: false,
                            ShowRepairs: false
                        );

                        // Convert to PDF
                        doc.SaveAs2(outputPath, FileFormat: 17);

                        result.Success = true;
                        result.OutputPath = outputPath;
                        
                        stopwatch.Stop();
                        _logger.Information("Converted {File} in {ElapsedMs}ms using PID {ProcessId}", 
                            Path.GetFileName(inputPath), stopwatch.ElapsedMilliseconds, wordInfo.ProcessId);
                    }
                    finally
                    {
                        // Clean up temporary file
                        if (tempFilePath != null && File.Exists(tempFilePath))
                        {
                            try { File.Delete(tempFilePath); } catch { }
                        }
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    _logger.Error(comEx, "COM error during conversion: {HResult}", comEx.HResult);
                    result.Success = false;
                    result.ErrorMessage = $"COM error: {comEx.Message}";
                    wordInfo.IsHealthy = false; // Mark as unhealthy
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to convert Word document: {Path}", inputPath);
                    result.Success = false;
                    result.ErrorMessage = $"Conversion failed: {ex.Message}";
                    wordInfo.IsHealthy = false; // Mark as unhealthy
                }
                finally
                {
                    // Clean up document
                    if (doc != null)
                    {
                        try
                        {
                            doc.Close(SaveChanges: false);
                            ComHelper.SafeReleaseComObject(doc, "Document", "RobustConversion");
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error closing Word document");
                        }
                    }
                }

                return result;
            });
        }

        private bool IsNetworkPath(string path)
        {
            try
            {
                return path.StartsWith(@"\\") || 
                       (Path.IsPathRooted(path) && new DriveInfo(path).DriveType == DriveType.Network);
            }
            catch
            {
                return false;
            }
        }

        private void PerformCleanup(object? state)
        {
            lock (_cleanupLock)
            {
                if (_conversionsCount > 3)
                {
                    _logger.Debug("Performing cleanup after {Count} conversions", _conversionsCount);
                    
                    // Clear pool of old instances
                    var instancesToDispose = new List<WordAppInfo>();
                    while (_wordAppPool.TryDequeue(out var instance))
                    {
                        if (DateTime.UtcNow - instance.CreatedAt > TimeSpan.FromMinutes(30))
                        {
                            instancesToDispose.Add(instance);
                        }
                        else
                        {
                            _wordAppPool.Enqueue(instance); // Keep recent ones
                        }
                    }
                    
                    // Dispose old instances
                    foreach (var instance in instancesToDispose)
                    {
                        _ = DisposeWordAppInfo(instance);
                    }
                    
                    // Memory cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    
                    _conversionsCount = 0;
                }
            }
        }

        private async Task DisposeWordAppInfo(WordAppInfo wordInfo)
        {
            try
            {
                if (wordInfo.Application != null)
                {
                    await Task.Run(() =>
                    {
                        try
                        {
                            wordInfo.Application.Quit();
                            ComHelper.SafeReleaseComObject(wordInfo.Application, "WordApp", "RobustDisposeWordApp");
                            _logger.Debug("Disposed Word application PID {ProcessId}", wordInfo.ProcessId);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Error disposing Word application");
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to dispose Word application PID {ProcessId}", wordInfo.ProcessId);
            }
        }

        public bool IsOfficeInstalled()
        {
            return IsOfficeAvailable();
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _cleanupTimer?.Dispose();
                _conversionSemaphore?.Dispose();
                
                // Dispose all pooled Word applications
                while (_wordAppPool.TryDequeue(out var instance))
                {
                    _ = DisposeWordAppInfo(instance);
                }
                
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                _disposed = true;
                _logger.Information("Robust Office conversion service disposed");
            }
        }
    }

    public class WordAppInfo
    {
        public dynamic? Application { get; set; }
        public int ProcessId { get; set; }
        public int ThreadId { get; set; }
        public DateTime CreatedAt { get; set; }
        public bool IsHealthy { get; set; }
    }
} 