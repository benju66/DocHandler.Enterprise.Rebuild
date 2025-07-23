using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Serilog;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace DocHandler.Services
{
    public class PdfCacheService : IDisposable
    {
        private readonly ILogger _logger;
        private readonly ConcurrentDictionary<string, CachedPdf> _cache;
        private readonly Timer _cleanupTimer;
        private readonly Timer _memoryMonitorTimer;
        private readonly TimeSpan _cacheExpiration = TimeSpan.FromMinutes(30);
        private readonly string _cacheDirectory;
        private readonly long _maxCacheSizeBytes;
        private readonly int _maxCacheEntries;
        private readonly SemaphoreSlim _cleanupSemaphore = new(1, 1);
        private long _currentCacheSizeBytes;
        private bool _disposed;
        
        public class CachedPdf
        {
            public string OriginalPath { get; set; } = "";
            public string CachedPath { get; set; } = "";
            public DateTime CreatedAt { get; set; }
            public DateTime LastAccessed { get; set; }
            public long FileSize { get; set; }
            public string FileHash { get; set; } = "";
            public int AccessCount { get; set; }
        }
        
        public class CacheStatistics
        {
            public int TotalEntries { get; set; }
            public long TotalSizeBytes { get; set; }
            public int HitCount { get; set; }
            public int MissCount { get; set; }
            public double HitRatio => TotalRequests > 0 ? (double)HitCount / TotalRequests : 0;
            public int TotalRequests => HitCount + MissCount;
            public string FormattedSize => FormatBytes(TotalSizeBytes);
            
            private static string FormatBytes(long bytes)
            {
                if (bytes < 1024) return $"{bytes} B";
                if (bytes < 1024 * 1024) return $"{bytes / 1024:F1} KB";
                if (bytes < 1024 * 1024 * 1024) return $"{bytes / (1024 * 1024):F1} MB";
                return $"{bytes / (1024 * 1024 * 1024):F1} GB";
            }
        }
        
        private int _hitCount;
        private int _missCount;
        
        public PdfCacheService(long maxCacheSizeMB = 500, int maxCacheEntries = 1000)
        {
            _logger = Log.ForContext<PdfCacheService>();
            _cache = new ConcurrentDictionary<string, CachedPdf>();
            _maxCacheSizeBytes = maxCacheSizeMB * 1024 * 1024;
            _maxCacheEntries = maxCacheEntries;
            
            // Create cache directory
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _cacheDirectory = Path.Combine(appData, "DocHandler", "PdfCache");
            
            try
            {
                Directory.CreateDirectory(_cacheDirectory);
                _logger.Information("PDF cache service initialized at {Path} with max size {MaxSize}MB", 
                    _cacheDirectory, maxCacheSizeMB);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to create cache directory");
                throw;
            }
            
            // Start cleanup timer (every 5 minutes)
            _cleanupTimer = new Timer(CleanupExpiredCache, null, 
                TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
            
            // Start memory monitor timer (every minute)
            _memoryMonitorTimer = new Timer(MonitorMemoryPressure, null,
                TimeSpan.FromMinutes(1), TimeSpan.FromMinutes(1));
        }
        
        public async Task<string?> GetCachedPdfAsync(string originalPath, string fileHash)
        {
            var cacheKey = GetCacheKey(originalPath, fileHash);
            
            if (_cache.TryGetValue(cacheKey, out var cached))
            {
                if (File.Exists(cached.CachedPath))
                {
                    // Update access statistics
                    cached.LastAccessed = DateTime.UtcNow;
                    cached.AccessCount++;
                    Interlocked.Increment(ref _hitCount);
                    
                    _logger.Debug("PDF cache hit for {File}", Path.GetFileName(originalPath));
                    return cached.CachedPath;
                }
                else
                {
                    // Cache file missing, remove entry and update size
                    if (_cache.TryRemove(cacheKey, out var removed))
                    {
                        Interlocked.Add(ref _currentCacheSizeBytes, -removed.FileSize);
                    }
                }
            }
            
            Interlocked.Increment(ref _missCount);
            return null;
        }
        
        public async Task<string> AddToCacheAsync(string originalPath, string pdfPath, string fileHash)
        {
            if (_disposed) return pdfPath;
            
            var cacheKey = GetCacheKey(originalPath, fileHash);
            var cachedFileName = $"{fileHash}_{Guid.NewGuid():N}_{Path.GetFileNameWithoutExtension(originalPath)}.pdf";
            var cachedPath = Path.Combine(_cacheDirectory, cachedFileName);
            
            try
            {
                var fileInfo = new FileInfo(pdfPath);
                var fileSize = fileInfo.Length;
                
                // Check if adding this file would exceed limits
                if (fileSize > _maxCacheSizeBytes / 10) // Don't cache files larger than 10% of total cache
                {
                    _logger.Debug("File too large for cache: {File} ({Size} bytes)", 
                        Path.GetFileName(originalPath), fileSize);
                    return pdfPath;
                }
                
                // Ensure we have space for the new file
                await EnsureCacheSpaceAsync(fileSize);
                
                // Copy to cache
                await Task.Run(() => File.Copy(pdfPath, cachedPath, true));
                
                var cached = new CachedPdf
                {
                    OriginalPath = originalPath,
                    CachedPath = cachedPath,
                    CreatedAt = DateTime.UtcNow,
                    LastAccessed = DateTime.UtcNow,
                    FileSize = fileSize,
                    FileHash = fileHash,
                    AccessCount = 0
                };
                
                _cache[cacheKey] = cached;
                Interlocked.Add(ref _currentCacheSizeBytes, fileSize);
                
                _logger.Debug("Added PDF to cache: {File} ({Size} bytes)", 
                    Path.GetFileName(originalPath), fileSize);
                
                return cachedPath;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to cache PDF for {File}", originalPath);
                
                // Clean up partial file
                try
                {
                    if (File.Exists(cachedPath))
                        File.Delete(cachedPath);
                }
                catch { /* Ignore cleanup errors */ }
                
                return pdfPath;
            }
        }
        
        private async Task EnsureCacheSpaceAsync(long requiredBytes)
        {
            if (_disposed) return;
            
            // Quick check without lock
            if (_currentCacheSizeBytes + requiredBytes <= _maxCacheSizeBytes && 
                _cache.Count < _maxCacheEntries)
            {
                return;
            }
            
            await _cleanupSemaphore.WaitAsync();
            try
            {
                // Double-check after acquiring lock
                while ((_currentCacheSizeBytes + requiredBytes > _maxCacheSizeBytes || 
                       _cache.Count >= _maxCacheEntries) && _cache.Count > 0)
                {
                    await EvictLeastRecentlyUsedAsync();
                }
            }
            finally
            {
                _cleanupSemaphore.Release();
            }
        }
        
        private async Task EvictLeastRecentlyUsedAsync()
        {
            if (_cache.IsEmpty) return;
            
            // Find LRU entry (least recently accessed with lowest access count as tiebreaker)
            var lruEntry = _cache.Values
                .OrderBy(x => x.LastAccessed)
                .ThenBy(x => x.AccessCount)
                .FirstOrDefault();
            
            if (lruEntry != null)
            {
                var cacheKey = GetCacheKey(lruEntry.OriginalPath, lruEntry.FileHash);
                if (_cache.TryRemove(cacheKey, out var removed))
                {
                    try
                    {
                        await Task.Run(() =>
                        {
                            if (File.Exists(removed.CachedPath))
                                File.Delete(removed.CachedPath);
                        });
                        
                        Interlocked.Add(ref _currentCacheSizeBytes, -removed.FileSize);
                        _logger.Debug("Evicted LRU cache entry: {File}", Path.GetFileName(removed.OriginalPath));
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to delete evicted cache file: {Path}", removed.CachedPath);
                    }
                }
            }
        }
        
        public CacheStatistics GetStatistics()
        {
            return new CacheStatistics
            {
                TotalEntries = _cache.Count,
                TotalSizeBytes = _currentCacheSizeBytes,
                HitCount = _hitCount,
                MissCount = _missCount
            };
        }
        
        public async Task ClearCacheAsync()
        {
            await _cleanupSemaphore.WaitAsync();
            try
            {
                _logger.Information("Clearing PDF cache");
                
                var entries = _cache.Values.ToList();
                _cache.Clear();
                _currentCacheSizeBytes = 0;
                
                // Delete files in background
                _ = Task.Run(async () =>
                {
                    foreach (var entry in entries)
                    {
                        try
                        {
                            if (File.Exists(entry.CachedPath))
                                File.Delete(entry.CachedPath);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Failed to delete cache file during clear: {Path}", entry.CachedPath);
                        }
                    }
                });
            }
            finally
            {
                _cleanupSemaphore.Release();
            }
        }
        
        private string GetCacheKey(string originalPath, string fileHash)
        {
            return $"{fileHash}_{Path.GetFileName(originalPath)}";
        }
        
        private async void CleanupExpiredCache(object? state)
        {
            if (_disposed || !_cleanupSemaphore.Wait(100)) return;
            
            try
            {
                var expiredEntries = _cache.Values
                    .Where(x => DateTime.UtcNow - x.LastAccessed > _cacheExpiration)
                    .ToList();
                
                if (expiredEntries.Any())
                {
                    _logger.Debug("Cleaning up {Count} expired cache entries", expiredEntries.Count);
                    
                    foreach (var entry in expiredEntries)
                    {
                        var cacheKey = GetCacheKey(entry.OriginalPath, entry.FileHash);
                        if (_cache.TryRemove(cacheKey, out var removed))
                        {
                            try
                            {
                                if (File.Exists(removed.CachedPath))
                                    File.Delete(removed.CachedPath);
                                
                                Interlocked.Add(ref _currentCacheSizeBytes, -removed.FileSize);
                            }
                            catch (Exception ex)
                            {
                                _logger.Warning(ex, "Failed to delete expired cache file: {Path}", removed.CachedPath);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error during cache cleanup");
            }
            finally
            {
                _cleanupSemaphore.Release();
            }
        }
        
        private async void MonitorMemoryPressure(object? state)
        {
            if (_disposed) return;
            
            try
            {
                var stats = GetStatistics();
                
                // Log statistics periodically
                if (stats.TotalEntries > 0)
                {
                    _logger.Debug("Cache stats: {Entries} entries, {Size}, {HitRatio:P1} hit ratio", 
                        stats.TotalEntries, stats.FormattedSize, stats.HitRatio);
                }
                
                // Check for memory pressure
                var memoryPressure = GC.GetTotalMemory(false);
                if (memoryPressure > 500 * 1024 * 1024) // 500MB threshold
                {
                    _logger.Information("High memory pressure detected, reducing cache size");
                    await EnsureCacheSpaceAsync(_maxCacheSizeBytes / 4); // Reduce to 75% of max
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during memory pressure monitoring");
            }
        }
        
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            
            _logger.Information("Disposing PDF cache service");
            
            _cleanupTimer?.Dispose();
            _memoryMonitorTimer?.Dispose();
            _cleanupSemaphore?.Dispose();
            
            // Final statistics
            var stats = GetStatistics();
            _logger.Information("Final cache statistics: {Entries} entries, {Size}, {HitRatio:P1} hit ratio", 
                stats.TotalEntries, stats.FormattedSize, stats.HitRatio);
        }
    }
} 