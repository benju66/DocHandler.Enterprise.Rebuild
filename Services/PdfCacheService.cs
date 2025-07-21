using System;
using System.Collections.Concurrent;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Serilog;
using System.Collections.Generic; // Added missing import

namespace DocHandler.Services
{
    public class PdfCacheService : IPdfCacheService
    {
        private readonly ILogger _logger;
        private readonly ConcurrentDictionary<string, CachedPdf> _cache;
        private readonly Timer _cleanupTimer;
        private readonly TimeSpan _cacheExpiration = TimeSpan.FromMinutes(30);
        private readonly string _cacheDirectory;
        private bool _disposed;
        
        public class CachedPdf
        {
            public string OriginalPath { get; set; } = "";
            public string CachedPath { get; set; } = "";
            public DateTime CreatedAt { get; set; }
            public DateTime LastAccessed { get; set; }
            public long FileSize { get; set; }
            public string FileHash { get; set; } = "";
        }
        
        public PdfCacheService()
        {
            _logger = Log.ForContext<PdfCacheService>();
            _cache = new ConcurrentDictionary<string, CachedPdf>();
            
            // Create cache directory asynchronously to avoid blocking
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            _cacheDirectory = Path.Combine(appData, "DocHandler", "PdfCache");
            
            // Create directory asynchronously
            _ = Task.Run(() => Directory.CreateDirectory(_cacheDirectory));
            
            // Start cleanup timer
            _cleanupTimer = new Timer(CleanupExpiredCache, null, 
                TimeSpan.FromMinutes(5), TimeSpan.FromMinutes(5));
            
            _logger.Information("PDF cache service initialized at {Path}", _cacheDirectory);
        }
        
        public async Task<string?> GetCachedPdfAsync(string originalPath, string fileHash)
        {
            var cacheKey = GetCacheKey(originalPath, fileHash);
            
            if (_cache.TryGetValue(cacheKey, out var cached))
            {
                if (File.Exists(cached.CachedPath))
                {
                    cached.LastAccessed = DateTime.UtcNow;
                    _logger.Debug("PDF cache hit for {File}", Path.GetFileName(originalPath));
                    return cached.CachedPath;
                }
                else
                {
                    // Cache file missing, remove entry
                    _cache.TryRemove(cacheKey, out _);
                }
            }
            
            return null;
        }
        
        public async Task<string> AddToCacheAsync(string originalPath, string pdfPath, string fileHash)
        {
            var cacheKey = GetCacheKey(originalPath, fileHash);
            var cachedFileName = $"{fileHash}_{Path.GetFileNameWithoutExtension(originalPath)}.pdf";
            var cachedPath = Path.Combine(_cacheDirectory, cachedFileName);
            
            try
            {
                // Copy to cache
                await Task.Run(() => File.Copy(pdfPath, cachedPath, true));
                
                var cached = new CachedPdf
                {
                    OriginalPath = originalPath,
                    CachedPath = cachedPath,
                    CreatedAt = DateTime.UtcNow,
                    LastAccessed = DateTime.UtcNow,
                    FileSize = new FileInfo(cachedPath).Length,
                    FileHash = fileHash
                };
                
                _cache[cacheKey] = cached;
                _logger.Debug("Added PDF to cache: {File}", Path.GetFileName(originalPath));
                
                return cachedPath;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to cache PDF for {File}", originalPath);
                return pdfPath; // Return original if caching fails
            }
        }
        
        private string GetCacheKey(string path, string hash)
        {
            return $"{hash}_{Path.GetFileName(path).ToLowerInvariant()}";
        }
        
        private void CleanupExpiredCache(object? state)
        {
            try
            {
                var now = DateTime.UtcNow;
                var expiredKeys = new List<string>();
                
                foreach (var kvp in _cache)
                {
                    if (now - kvp.Value.LastAccessed > _cacheExpiration)
                    {
                        expiredKeys.Add(kvp.Key);
                    }
                }
                
                foreach (var key in expiredKeys)
                {
                    if (_cache.TryRemove(key, out var cached))
                    {
                        try
                        {
                            File.Delete(cached.CachedPath);
                            _logger.Debug("Removed expired cache: {File}", 
                                Path.GetFileName(cached.OriginalPath));
                        }
                        catch { }
                    }
                }
                
                if (expiredKeys.Count > 0)
                {
                    _logger.Information("Cleaned up {Count} expired cache entries", expiredKeys.Count);
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during cache cleanup");
            }
        }
        
        public void ClearCache()
        {
            foreach (var cached in _cache.Values)
            {
                try { File.Delete(cached.CachedPath); } catch { }
            }
            _cache.Clear();
            
            _logger.Information("PDF cache cleared");
        }
        
        public void Dispose()
        {
            if (!_disposed)
            {
                _cleanupTimer?.Dispose();
                ClearCache();
                _disposed = true;
            }
        }
    }
} 