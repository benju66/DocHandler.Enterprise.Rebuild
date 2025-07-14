using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Text;
using System.Security.Cryptography;
using System.Diagnostics;
using Serilog;
using Task = System.Threading.Tasks.Task;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DocHandler.Services
{
    public class CompanyDetectionSettings
    {
        public int MaxPagesForFullScan { get; set; } = 5;
        public int MaxCharactersToExtract { get; set; } = 10000;
        public bool EnableOCR { get; set; } = false; // Disabled for now
        public bool EnableFuzzyMatching { get; set; } = true;
        public double MinimumMatchScore { get; set; } = 0.7;
        public int CacheExpiryMinutes { get; set; } = 30;
        public bool EnableParallelProcessing { get; set; } = true;
        public int MaxConcurrentOperations { get; set; } = Environment.ProcessorCount;
        
        // Safe performance optimizations
        public int MaxCacheSize { get; set; } = 200; // Increased cache size
        public int MemoryCleanupThreshold { get; set; } = 100; // MB - prevents memory exhaustion
        public bool EnableMemoryMonitoring { get; set; } = true; // Helps prevent DoS attacks
        public bool EnablePerformanceMetrics { get; set; } = true; // Safe performance tracking
    }
    
    public class PerformanceMetrics
    {
        public int DocumentsProcessed { get; set; }
        public TimeSpan TotalProcessingTime { get; set; }
        public int CacheHits { get; set; }
        public int CacheMisses { get; set; }
        public long PeakMemoryUsage { get; set; }
        public DateTime LastReset { get; set; } = DateTime.Now;
        
        public double AverageProcessingTime => 
            DocumentsProcessed > 0 ? TotalProcessingTime.TotalMilliseconds / DocumentsProcessed : 0;
            
        public double CacheHitRate => 
            (CacheHits + CacheMisses) > 0 ? (double)CacheHits / (CacheHits + CacheMisses) * 100 : 0;
            
        public string GetSummary()
        {
            return $"Processed: {DocumentsProcessed}, Avg Time: {AverageProcessingTime:F1}ms, " +
                   $"Cache Hit Rate: {CacheHitRate:F1}%, Peak Memory: {PeakMemoryUsage / 1024 / 1024}MB";
        }
    }
    
    public class MatchResult
    {
        public double Score { get; set; }
        public string Method { get; set; } = "";
    }
    
    public class CompanyDetectionResult
    {
        public string FilePath { get; set; } = "";
        public string? DetectedCompany { get; set; }
        public DateTime ProcessedAt { get; set; }
        public double ConfidenceScore { get; set; }
    }

    public class CompanyNameService
    {
        private readonly ILogger _logger;
        private readonly string _dataPath;
        private readonly string _companyNamesPath;
        private CompanyNamesData _data;
        private readonly ConcurrentDictionary<string, string> _textCache;
        private readonly ConcurrentDictionary<string, (string?, DateTime)> _detectionCache;
        private readonly CompanyDetectionSettings _settings;
        private readonly PerformanceMetrics _performanceMetrics;
        private DateTime _lastMemoryCheck = DateTime.Now;
        
        // Performance metrics
        private int _detectionCount = 0;
        private readonly object _metricsLock = new object();
        
        // PDF caching for avoiding double conversion
        private readonly ConcurrentDictionary<string, string> _convertedPdfCache = new();
        private readonly ConcurrentDictionary<string, DateTime> _pdfCacheTimestamps = new();
        private readonly ConcurrentDictionary<string, FileInfo> _pdfCacheFileInfo = new();
        private readonly TimeSpan _cacheExpiration = TimeSpan.FromMinutes(30);
        private Timer _cacheCleanupTimer;
        private readonly object _cacheCleanupLock = new object();

        // Services for conversion
        private SessionAwareOfficeService? _sessionOfficeService;
        private SessionAwareExcelService? _sessionExcelService;
        
        // Configurable .doc file size limit
        private int _docFileSizeLimitMB = 10;
        
        // Safe compiled regex patterns for better performance
        private static readonly Regex ControlCharactersRegex = new(@"[\x00-\x1F\x7F-\x9F]", RegexOptions.Compiled);
        private static readonly Regex MultipleSpacesRegex = new(@"\s+", RegexOptions.Compiled);
        private static readonly Regex NonWordCharactersRegex = new(@"[^\w\s.,&\-()]", RegexOptions.Compiled);
        
        public List<CompanyInfo> Companies => _data.Companies;
        
        public CompanyNameService()
        {
            _logger = Log.ForContext<CompanyNameService>();
            
            // Initialize thread-safe caching with concurrent dictionaries
            _textCache = new ConcurrentDictionary<string, string>();
            _detectionCache = new ConcurrentDictionary<string, (string?, DateTime)>();
            
            // Load performance settings
            _settings = new CompanyDetectionSettings();
            
            // Initialize performance metrics
            _performanceMetrics = new PerformanceMetrics();
            
            _logger.Information("Enhanced company detection service initialized with caching, fuzzy matching, and performance monitoring");
            
            // Store data in AppData
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "DocHandler");
            Directory.CreateDirectory(appFolder);
            
            _dataPath = appFolder;
            _companyNamesPath = Path.Combine(appFolder, "company_names.json");
            _data = LoadCompanyNames();
            
            // Start periodic cache cleanup every 15 minutes
            _cacheCleanupTimer = new Timer(
                _ => CleanupPdfCache(), 
                null, 
                TimeSpan.FromMinutes(15), 
                TimeSpan.FromMinutes(15)
            );
            
            _logger.Information("PDF caching initialized with {Expiration} minute expiration", _cacheExpiration.TotalMinutes);
        }
        
        // Injected services setter
        public void SetOfficeServices(SessionAwareOfficeService officeService, SessionAwareExcelService excelService)
        {
            _sessionOfficeService = officeService;
            _sessionExcelService = excelService;
        }
        
        public void UpdateDocFileSizeLimit(int limitMB)
        {
            _docFileSizeLimitMB = limitMB;
            _logger.Information("Updated .doc file size limit to {LimitMB}MB", limitMB);
        }
        
        private CompanyNamesData LoadCompanyNames()
        {
            try
            {
                if (File.Exists(_companyNamesPath))
                {
                    var json = File.ReadAllText(_companyNamesPath);
                    var options = new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        AllowTrailingCommas = true,
                        ReadCommentHandling = JsonCommentHandling.Skip
                    };
                    var data = JsonSerializer.Deserialize<CompanyNamesData>(json, options);
                    
                    if (data != null && data.Companies != null)
                    {
                        _logger.Information("Loaded {Count} company names", data.Companies.Count);
                        return data;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to load company names");
            }
            
            // Return default data with some sample companies
            _logger.Information("Creating default company names data");
            var defaultData = CreateDefaultData();
            
            // Save the default data immediately
            _ = SaveCompanyNames();
            
            return defaultData;
        }
        
        private CompanyNamesData CreateDefaultData()
        {
            return new CompanyNamesData
            {
                Companies = new List<CompanyInfo>
                {
                    new CompanyInfo { Name = "ABC Construction", Aliases = new List<string> { "ABC", "ABC Const" } },
                    new CompanyInfo { Name = "Smith Electrical", Aliases = new List<string> { "Smith Electric", "Smith" } },
                    new CompanyInfo { Name = "Johnson Plumbing", Aliases = new List<string> { "Johnson", "Johnson Plumb" } },
                    new CompanyInfo { Name = "XYZ Contractors", Aliases = new List<string> { "XYZ" } }
                }
            };
        }
        
        public async Task SaveCompanyNames()
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true
                };
                
                var json = JsonSerializer.Serialize(_data, options);
                await File.WriteAllTextAsync(_companyNamesPath, json).ConfigureAwait(false);
                
                _logger.Information("Company names saved to {Path}", _companyNamesPath);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to save company names");
            }
        }
        
        public async Task<bool> AddCompanyName(string name, List<string>? aliases = null)
        {
            try
            {
                name = name.Trim();
                
                // Check if already exists
                if (_data.Companies.Any(c => c.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                {
                    _logger.Warning("Company name already exists: {Name}", name);
                    return false;
                }
                
                var company = new CompanyInfo
                {
                    Name = name,
                    Aliases = aliases ?? new List<string>(),
                    DateAdded = DateTime.Now,
                    UsageCount = 0
                };
                
                _data.Companies.Add(company);
                _data.Companies = _data.Companies.OrderBy(c => c.Name).ToList();
                
                await SaveCompanyNames().ConfigureAwait(false);
                _logger.Information("Added new company: {Name}", name);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to add company name: {Name}", name);
                return false;
            }
        }
        
        public async Task<bool> RemoveCompanyName(string name)
        {
            try
            {
                var company = _data.Companies.FirstOrDefault(c => 
                    c.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
                
                if (company != null)
                {
                    _data.Companies.Remove(company);
                    await SaveCompanyNames().ConfigureAwait(false);
                    _logger.Information("Removed company: {Name}", name);
                    return true;
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to remove company name: {Name}", name);
                return false;
            }
        }
        
        public async Task<string?> ScanDocumentForCompanyName(string filePath, IProgress<int>? progress = null)
        {
            var stopwatch = System.Diagnostics.Stopwatch.StartNew();
            Interlocked.Increment(ref _detectionCount);
            bool cachedDetection = false;
            
            try
            {
                progress?.Report(10); // Starting scan
                
                // Check cache first
                var fileInfo = new FileInfo(filePath);
                if (!fileInfo.Exists)
                {
                    _logger.Warning("File does not exist: {Path}", filePath);
                    
                    stopwatch.Stop();
                    lock (_metricsLock)
                    {
                        _performanceMetrics.DocumentsProcessed++;
                        _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                        _performanceMetrics.CacheMisses++;
                    }
                    return null;
                }
                
                // Check if this is a .doc file and handle specially
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                if (extension == ".doc")
                {
                    _logger.Information("Processing .doc file for company detection: {Path}", filePath);
                    
                    // Add timeout and resource management for .doc files
                    using var docCts = new CancellationTokenSource(TimeSpan.FromSeconds(15)); // Shorter timeout for .doc files
                    
                    try
                    {
                        return await ProcessDocFileForCompanyDetection(filePath, progress, docCts.Token);
                    }
                    catch (OperationCanceledException)
                    {
                        _logger.Warning("Company detection timed out for .doc file: {Path}", filePath);
                        return null;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to process .doc file for company detection: {Path}", filePath);
                        return null;
                    }
                }
                
                // Early exit for very large files
                var fileSizeMB = fileInfo.Length / (1024.0 * 1024.0);
                if (fileSizeMB > 50) // Skip files larger than 50MB
                {
                    _logger.Warning("Skipping very large file ({Size:F1}MB): {Path}", fileSizeMB, filePath);
                    return null;
                }
                
                var cacheKey = $"{filePath}_{fileInfo.LastWriteTime.Ticks}_{fileInfo.Length}";
                
                progress?.Report(20); // Checking cache
                
                // Check detection cache
                if (_detectionCache.TryGetValue(cacheKey, out var cachedDetectionValue))
                {
                    var cacheAge = DateTime.Now - cachedDetectionValue.Item2;
                    if (cacheAge.TotalMinutes < _settings.CacheExpiryMinutes)
                    {
                        _logger.Debug("Using cached detection result for {File}", Path.GetFileName(filePath));
                        cachedDetection = true;
                        
                        progress?.Report(100); // Complete from cache
                        
                        stopwatch.Stop();
                        lock (_metricsLock)
                        {
                            _performanceMetrics.DocumentsProcessed++;
                            _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                            _performanceMetrics.CacheHits++;
                        }
                        return cachedDetectionValue.Item1;
                    }
                    else
                    {
                        _detectionCache.TryRemove(cacheKey, out _);
                    }
                }
                
                _logger.Information("Scanning document for company names: {Path}", filePath);
                
                progress?.Report(30); // Starting text extraction
                
                // Get cached or extract text
                string documentText = await GetCachedDocumentText(filePath, cacheKey, progress).ConfigureAwait(false);
                
                progress?.Report(60); // Text extraction complete
                
                if (string.IsNullOrWhiteSpace(documentText))
                {
                    _logger.Warning("No text extracted from document: {Path}", filePath);
                    _detectionCache[cacheKey] = (null, DateTime.Now);
                    
                    stopwatch.Stop();
                    lock (_metricsLock)
                    {
                        _performanceMetrics.DocumentsProcessed++;
                        _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                        _performanceMetrics.CacheMisses++;
                    }
                    return null;
                }
                
                // Validate that we have meaningful text
                if (documentText.Trim().Length < 10)
                {
                    _logger.Warning("Insufficient text content for company detection: {Length} characters", documentText.Trim().Length);
                    _detectionCache[cacheKey] = (null, DateTime.Now);
                    
                    stopwatch.Stop();
                    lock (_metricsLock)
                    {
                        _performanceMetrics.DocumentsProcessed++;
                        _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                        _performanceMetrics.CacheMisses++;
                    }
                    return null;
                }
                
                _logger.Debug("Extracted {Length} characters from document for company detection", documentText.Length);
                
                progress?.Report(70); // Starting company matching
                
                // Preprocess text for better matching
                string processedText = PreprocessText(documentText);
                
                // Check if we have any companies to search for
                if (!_data.Companies.Any())
                {
                    _logger.Warning("No companies in database to search for");
                    _detectionCache[cacheKey] = (null, DateTime.Now);
                    
                    stopwatch.Stop();
                    lock (_metricsLock)
                    {
                        _performanceMetrics.DocumentsProcessed++;
                        _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                        _performanceMetrics.CacheMisses++;
                    }
                    return null;
                }
                
                progress?.Report(80); // Finding best match
                
                // Find best company match using enhanced matching
                var bestMatch = await FindBestCompanyMatch(processedText).ConfigureAwait(false);
                
                progress?.Report(90); // Processing results
                
                if (bestMatch != null)
                {
                    _logger.Information("Found company: {Name} (Score: {Score})", bestMatch.Value.company, bestMatch.Value.score);
                    await IncrementUsageCount(bestMatch.Value.company).ConfigureAwait(false);
                    _detectionCache[cacheKey] = (bestMatch.Value.company, DateTime.Now);
                    
                    progress?.Report(100); // Complete
                    
                    stopwatch.Stop();
                    lock (_metricsLock)
                    {
                        _performanceMetrics.DocumentsProcessed++;
                        _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                        _performanceMetrics.CacheMisses++;
                    }
                    return bestMatch.Value.company;
                }
                
                _logger.Information("No known company names found in document (searched {Count} companies)", _data.Companies.Count);
                _detectionCache[cacheKey] = (null, DateTime.Now);
                
                progress?.Report(100); // Complete
                
                stopwatch.Stop();
                lock (_metricsLock)
                {
                    _performanceMetrics.DocumentsProcessed++;
                    _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                    _performanceMetrics.CacheMisses++;
                }
                return null;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to scan document for company names: {Path}", filePath);
                
                stopwatch.Stop();
                lock (_metricsLock)
                {
                    _performanceMetrics.DocumentsProcessed++;
                    _performanceMetrics.TotalProcessingTime = _performanceMetrics.TotalProcessingTime.Add(stopwatch.Elapsed);
                    _performanceMetrics.CacheMisses++;
                }
                return null;
            }
        }
        
        private async Task<string> GetCachedDocumentText(string filePath, string cacheKey, IProgress<int>? progress = null)
        {
            if (_textCache.TryGetValue(cacheKey, out string? cachedText))
            {
                _logger.Debug("Using cached text extraction for {File}", Path.GetFileName(filePath));
                progress?.Report(60); // Update progress for cached text
                return cachedText;
            }
            
            progress?.Report(40); // Starting text extraction
            
            var extractedText = await ExtractTextFromDocument(filePath, progress).ConfigureAwait(false);
            
            progress?.Report(55); // Text extraction complete, caching
            
            // Cache the extracted text (limit cache size)
            if (_textCache.Count > 100)
            {
                // Remove oldest entries
                var oldestKeys = _textCache.Keys.Take(_textCache.Count - 80).ToList();
                foreach (var key in oldestKeys)
                {
                    _textCache.TryRemove(key, out _);
                }
            }
            
            _textCache[cacheKey] = extractedText;
            return extractedText;
        }
        
        private string PreprocessText(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return text;
            
            // Remove common PDF artifacts and normalize text
            text = Regex.Replace(text, @"[\x00-\x1F\x7F-\x9F]", " "); // Control characters
            text = Regex.Replace(text, @"\s+", " "); // Multiple spaces
            text = Regex.Replace(text, @"[^\w\s.,&\-()]", " "); // Keep only useful punctuation
            
            return text.Trim().ToLowerInvariant();
        }
        
        private async Task<(string company, double score)?> FindBestCompanyMatch(string text)
        {
            var matches = new List<(string company, double score)>();
            
            // Check each company and its aliases
            foreach (var company in _data.Companies.OrderByDescending(c => c.UsageCount))
            {
                if (string.IsNullOrWhiteSpace(company.Name)) continue;
                
                // Check main name
                var mainScore = GetCompanyMatchScore(text, company.Name);
                if (mainScore > 0)
                {
                    matches.Add((company.Name, mainScore));
                }
                
                // Check aliases
                foreach (var alias in company.Aliases)
                {
                    if (!string.IsNullOrWhiteSpace(alias))
                    {
                        var aliasScore = GetCompanyMatchScore(text, alias) * 0.9; // Slightly lower score for aliases
                        if (aliasScore > 0)
                        {
                            matches.Add((company.Name, aliasScore));
                        }
                    }
                }
            }
            
            // Return best match above threshold
            var bestMatch = matches.OrderByDescending(m => m.score).FirstOrDefault();
            return bestMatch.score >= _settings.MinimumMatchScore ? bestMatch : null;
        }
        
        private double GetCompanyMatchScore(string text, string companyName)
        {
            if (companyName.Length < 2) return 0.0;
            
            var normalizedCompany = companyName.ToLowerInvariant().Trim();
            var matches = new List<MatchResult>();
            
            // 1. Exact word boundary match (highest priority)
            matches.Add(CheckExactWordMatch(text, normalizedCompany));
            
            // 2. Fuzzy matching for common variations
            if (_settings.EnableFuzzyMatching)
            {
                matches.Add(CheckFuzzyMatch(text, normalizedCompany));
            }
            
            // 3. Contextual matching
            matches.Add(CheckContextualMatch(text, normalizedCompany));
            
            // Return best score
            var bestMatch = matches.OrderByDescending(m => m.Score).FirstOrDefault();
            return bestMatch?.Score ?? 0.0;
        }
        
        private MatchResult CheckExactWordMatch(string text, string companyName)
        {
            try
            {
                var pattern = $@"\b{Regex.Escape(companyName)}\b";
                var match = Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
                return new MatchResult
                {
                    Score = match ? 1.0 : 0.0,
                    Method = "ExactWord"
                };
            }
            catch
            {
                return new MatchResult { Score = 0.0, Method = "ExactWord" };
            }
        }
        
        private MatchResult CheckFuzzyMatch(string text, string companyName)
        {
            var variations = GenerateCompanyVariations(companyName);
            
            foreach (var variation in variations)
            {
                if (text.Contains(variation))
                {
                    var score = CalculateSimilarityScore(companyName, variation);
                    return new MatchResult
                    {
                        Score = score * 0.9, // Slightly lower than exact match
                        Method = "Fuzzy"
                    };
                }
            }
            
            return new MatchResult { Score = 0.0, Method = "Fuzzy" };
        }
        
        private MatchResult CheckContextualMatch(string text, string companyName)
        {
            var contextPatterns = new[]
            {
                $@"from:\s*{Regex.Escape(companyName)}",
                $@"regards,?\s*{Regex.Escape(companyName)}",
                $@"sincerely,?\s*{Regex.Escape(companyName)}",
                $@"quote\s+from\s+{Regex.Escape(companyName)}",
                $@"{Regex.Escape(companyName)}\s+team",
                $@"contact\s+{Regex.Escape(companyName)}"
            };
            
            foreach (var pattern in contextPatterns)
            {
                if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                {
                    return new MatchResult
                    {
                        Score = 0.95,
                        Method = "Contextual"
                    };
                }
            }
            
            return new MatchResult { Score = 0.0, Method = "Contextual" };
        }
        
        private List<string> GenerateCompanyVariations(string companyName)
        {
            var variations = new List<string> { companyName };
            
            // Common abbreviations
            var abbreviations = new Dictionary<string, string>
            {
                { "corporation", "corp" },
                { "incorporated", "inc" },
                { "company", "co" },
                { "limited", "ltd" },
                { "and", "&" },
                { "construction", "const" },
                { "electrical", "electric" },
                { "plumbing", "plumb" }
            };
            
            foreach (var abbrev in abbreviations)
            {
                if (companyName.Contains(abbrev.Key))
                {
                    variations.Add(companyName.Replace(abbrev.Key, abbrev.Value));
                }
                if (companyName.Contains(abbrev.Value))
                {
                    variations.Add(companyName.Replace(abbrev.Value, abbrev.Key));
                }
            }
            
            // Remove punctuation variations
            variations.Add(Regex.Replace(companyName, @"[^\w\s]", ""));
            
            // Add spacing variations
            variations.Add(companyName.Replace(" ", ""));
            variations.Add(companyName.Replace("-", " "));
            variations.Add(companyName.Replace(".", " "));
            
            return variations.Distinct().ToList();
        }
        
        private double CalculateSimilarityScore(string original, string comparison)
        {
            // Simple Levenshtein distance-based similarity
            var distance = CalculateLevenshteinDistance(original, comparison);
            var maxLength = Math.Max(original.Length, comparison.Length);
            return maxLength == 0 ? 1.0 : 1.0 - (double)distance / maxLength;
        }
        
        private int CalculateLevenshteinDistance(string source, string target)
        {
            if (string.IsNullOrEmpty(source)) return target?.Length ?? 0;
            if (string.IsNullOrEmpty(target)) return source.Length;
            
            var sourceLength = source.Length;
            var targetLength = target.Length;
            var matrix = new int[sourceLength + 1, targetLength + 1];
            
            for (int i = 0; i <= sourceLength; matrix[i, 0] = i++) { }
            for (int j = 0; j <= targetLength; matrix[0, j] = j++) { }
            
            for (int i = 1; i <= sourceLength; i++)
            {
                for (int j = 1; j <= targetLength; j++)
                {
                    var cost = target[j - 1] == source[i - 1] ? 0 : 1;
                    matrix[i, j] = Math.Min(
                        Math.Min(matrix[i - 1, j] + 1, matrix[i, j - 1] + 1),
                        matrix[i - 1, j - 1] + cost);
                }
            }
            
            return matrix[sourceLength, targetLength];
        }
        
        private bool ContainsCompanyName(string text, string companyName)
        {
            try
            {
                // Skip very short company names (likely to cause false positives)
                if (companyName.Length < 3)
                {
                    _logger.Debug("Skipping very short company name: {Name}", companyName);
                    return false;
                }
                
                // Normalize the company name
                var normalizedCompanyName = companyName.ToLowerInvariant().Trim();
                
                // First try exact word boundary match
                string pattern = $@"\b{Regex.Escape(normalizedCompanyName)}\b";
                if (Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase))
                {
                    return true;
                }
                
                // For company names with special characters, also try without word boundaries
                // This handles cases like "ABC, Inc." or "XYZ Corp"
                if (normalizedCompanyName.Contains(",") || normalizedCompanyName.Contains(".") || 
                    normalizedCompanyName.Contains("&") || normalizedCompanyName.Contains("-"))
                {
                    return text.Contains(normalizedCompanyName);
                }
                
                return false;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Regex match failed for company name: {Name}", companyName);
                // Fallback to simple contains for normalized names only
                var normalizedCompanyName = companyName.ToLowerInvariant().Trim();
                if (normalizedCompanyName.Length >= 3)
                {
                    return text.Contains(normalizedCompanyName);
                }
                return false;
            }
        }
        
        private async Task<string> ExtractTextFromDocument(string filePath, IProgress<int>? progress = null)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            try
            {
                // For PDFs, extract directly (already fast)
                if (extension == ".pdf")
                {
                    progress?.Report(20);
                    return await ExtractTextFromPdf(filePath, progress);
                }
                
                // For Office documents, use fast OpenXML extraction first
                if (extension == ".docx" || extension == ".doc" || extension == ".xlsx" || extension == ".xls")
                {
                    progress?.Report(10);
                    
                    // Try fast OpenXML extraction first
                    string text = string.Empty;
                    
                    if (extension == ".docx" || extension == ".doc")
                    {
                        text = await ExtractTextFromWord(filePath);
                    }
                    else // Excel
                    {
                        text = await ExtractTextFromExcel(filePath);
                    }
                    
                    progress?.Report(80);
                    
                    // If OpenXML extraction was successful, return the text
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        _logger.Information("Extracted text via OpenXML for {File} - {Length} chars", 
                            Path.GetFileName(filePath), text.Length);
                        progress?.Report(100);
                        return text;
                    }
                    
                    // If OpenXML failed, fall back to PDF conversion approach
                    _logger.Warning("OpenXML extraction failed for {File}, falling back to PDF conversion", 
                        Path.GetFileName(filePath));
                    
                    // Check if we already have a converted PDF
                    if (_convertedPdfCache.TryGetValue(filePath, out string? cachedPdfPath) && 
                        File.Exists(cachedPdfPath))
                    {
                        _logger.Debug("Using cached PDF conversion for {File}", Path.GetFileName(filePath));
                        progress?.Report(50);
                        return await ExtractTextFromPdf(cachedPdfPath, progress);
                    }
                    
                    // Convert to temporary PDF as fallback
                    var tempPdf = Path.Combine(Path.GetTempPath(), $"DocHandler_{Guid.NewGuid()}.pdf");
                    ConversionResult conversionResult;
                    
                    if (extension == ".docx" || extension == ".doc")
                    {
                        if (_sessionOfficeService == null)
                        {
                            _logger.Warning("Office service not available for PDF conversion fallback");
                            return string.Empty;
                        }
                        
                        conversionResult = await _sessionOfficeService.ConvertWordToPdf(filePath, tempPdf);
                    }
                    else // Excel
                    {
                        if (_sessionExcelService == null)
                        {
                            _logger.Warning("Excel service not available for PDF conversion fallback");
                            return string.Empty;
                        }
                        
                        conversionResult = await _sessionExcelService.ConvertSpreadsheetToPdf(filePath, tempPdf);
                    }
                    
                    progress?.Report(40);
                    
                    if (!conversionResult.Success)
                    {
                        _logger.Warning("Failed to convert {File} to PDF: {Error}", 
                            Path.GetFileName(filePath), conversionResult.ErrorMessage);
                        return string.Empty;
                    }
                    
                    try
                    {
                        // Cache the converted PDF for reuse in final processing
                        _convertedPdfCache[filePath] = tempPdf;
                        _pdfCacheTimestamps[filePath] = DateTime.Now;
                        _pdfCacheFileInfo[filePath] = new FileInfo(filePath);
                        
                        // Extract text from the PDF
                        progress?.Report(50);
                        var pdfText = await ExtractTextFromPdf(tempPdf, progress);
                        
                        _logger.Information("Extracted text via PDF conversion fallback for {File} - {Length} chars", 
                            Path.GetFileName(filePath), pdfText.Length);
                        
                        return pdfText;
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Failed to extract text from converted PDF");
                        
                        // Clean up failed conversion
                        if (File.Exists(tempPdf))
                        {
                            try { File.Delete(tempPdf); } catch { }
                        }
                        _convertedPdfCache.TryRemove(filePath, out _);
                        
                        return string.Empty;
                    }
                }
                
                // Text files
                if (extension == ".txt")
                {
                    return await ExtractTextFromTextFile(filePath);
                }
                
                _logger.Warning("Unsupported file type for text extraction: {Extension}", extension);
                return string.Empty;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from document: {Path}", filePath);
                return string.Empty;
            }
        }
        
        private async Task<string> ExtractTextFromPdf(string filePath, IProgress<int>? progress = null)
        {
            try
            {
                using var pdfReader = new PdfReader(filePath);
                using var pdfDoc = new PdfDocument(pdfReader);
                
                var text = new StringBuilder();
                var strategy = new SimpleTextExtractionStrategy();
                
                // Only read first few pages for company detection (huge speedup!)
                int totalPages = pdfDoc.GetNumberOfPages();
                int pagesToRead = Math.Min(3, totalPages);
                
                _logger.Debug("Reading {Pages} of {Total} pages from PDF for company detection", 
                    pagesToRead, totalPages);
                
                for (int i = 1; i <= pagesToRead; i++)
                {
                    var page = pdfDoc.GetPage(i);
                    var pageText = PdfTextExtractor.GetTextFromPage(page, strategy);
                    text.AppendLine(pageText);
                    
                    // Update progress if provided
                    if (progress != null)
                    {
                        int progressValue = 50 + (i * 40 / pagesToRead);
                        progress.Report(progressValue);
                    }
                    
                    // Early termination if we have enough text (5KB should be plenty)
                    if (text.Length > 5000)
                    {
                        _logger.Debug("Early termination - sufficient text extracted ({Length} chars)", text.Length);
                        break;
                    }
                }
                
                progress?.Report(90);
                return text.ToString();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from PDF: {Path}", filePath);
                return string.Empty;
            }
        }
        
        private List<int> GetPriorityPages(int totalPages)
        {
            var pages = new List<int>();
            
            // Always scan first 2 pages (letterhead, headers)
            pages.AddRange(Enumerable.Range(1, Math.Min(2, totalPages)));
            
            // Always scan last 2 pages (signatures, company info)
            if (totalPages > 2)
            {
                var lastPagesStart = Math.Max(3, totalPages - 1);
                var lastPagesCount = Math.Min(2, totalPages - 2);
                if (lastPagesCount > 0)
                {
                    pages.AddRange(Enumerable.Range(lastPagesStart, lastPagesCount));
                }
            }
            
            // For medium documents, scan some middle pages
            if (totalPages > 4 && totalPages <= _settings.MaxPagesForFullScan)
            {
                var middleStart = 3;
                var middleEnd = Math.Min(totalPages - 2, _settings.MaxPagesForFullScan - 2);
                if (middleEnd > middleStart)
                {
                    pages.AddRange(Enumerable.Range(middleStart, middleEnd - middleStart + 1));
                }
            }
            
            // For very short documents, scan all pages
            if (totalPages <= 3)
            {
                pages.AddRange(Enumerable.Range(1, totalPages));
            }
            
            return pages.Distinct().OrderBy(p => p).ToList();
        }
        
        private async Task<string> ExtractTextFromWordDirect(string filePath)
        {
            return await Task.Run(() =>
            {
                var extractedText = new StringBuilder();
                var stopwatch = Stopwatch.StartNew();
                const int maxChars = 5000;
                const int timeoutMs = 2000;
                
                try
                {
                    // Check file extension - only process .docx files with OpenXML
                    var extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".docx")
                    {
                        _logger.Debug("Skipping OpenXML extraction for non-docx file: {Extension}", extension);
                        return string.Empty;
                    }
                    
                    using (var doc = WordprocessingDocument.Open(filePath, false))
                    {
                        // 1. Extract document properties first (highest priority)
                        try
                        {
                            var extendedProps = doc.ExtendedFilePropertiesPart;
                            if (extendedProps != null)
                            {
                                var props = extendedProps.Properties;
                                if (props != null)
                                {
                                    var company = props.Company?.Text;
                                    
                                    if (!string.IsNullOrWhiteSpace(company))
                                    {
                                        extractedText.AppendLine($"COMPANY: {company}");
                                        _logger.Debug("Found company in document properties: {Company}", company);
                                    }
                                }
                            }
                            
                            var coreProps = doc.PackageProperties;
                            if (coreProps != null)
                            {
                                var creator = coreProps.Creator;
                                
                                if (!string.IsNullOrWhiteSpace(creator))
                                {
                                    extractedText.AppendLine($"AUTHOR: {creator}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Debug(ex, "Failed to extract document properties");
                        }
                        
                        // Early termination check
                        if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                        {
                            _logger.Debug("Early termination after document properties");
                            return extractedText.ToString();
                        }
                        
                        // 2. Extract headers and footers (second priority)
                        try
                        {
                            foreach (var headerPart in doc.MainDocumentPart?.HeaderParts ?? Enumerable.Empty<HeaderPart>())
                            {
                                var headerText = headerPart.Header?.InnerText?.Trim();
                                if (!string.IsNullOrWhiteSpace(headerText))
                                {
                                    extractedText.AppendLine($"HEADER: {headerText}");
                                    
                                    if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                                        break;
                                }
                            }
                            
                            foreach (var footerPart in doc.MainDocumentPart?.FooterParts ?? Enumerable.Empty<FooterPart>())
                            {
                                var footerText = footerPart.Footer?.InnerText?.Trim();
                                if (!string.IsNullOrWhiteSpace(footerText))
                                {
                                    extractedText.AppendLine($"FOOTER: {footerText}");
                                    
                                    if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                                        break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Debug(ex, "Failed to extract headers/footers");
                        }
                        
                        // Early termination check
                        if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                        {
                            _logger.Debug("Early termination after headers/footers");
                            return extractedText.ToString();
                        }
                        
                        var body = doc.MainDocumentPart?.Document?.Body;
                        if (body == null)
                        {
                            _logger.Warning("Word document has no body content: {Path}", filePath);
                            return extractedText.ToString();
                        }
                        
                        var paragraphs = body.Elements<Paragraph>().ToList();
                        
                        // 3. Extract first 5 paragraphs (third priority)
                        var firstParagraphs = paragraphs.Take(5);
                        foreach (var para in firstParagraphs)
                        {
                            var paraText = para.InnerText?.Trim();
                            if (!string.IsNullOrWhiteSpace(paraText))
                            {
                                extractedText.AppendLine(paraText);
                                
                                if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                                    break;
                            }
                        }
                        
                        // Early termination check
                        if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                        {
                            _logger.Debug("Early termination after first 5 paragraphs");
                            return extractedText.ToString();
                        }
                        
                        // 4. Extract last 5 paragraphs (signature area - fourth priority)
                        if (paragraphs.Count > 5)
                        {
                            var lastParagraphs = paragraphs.Skip(Math.Max(0, paragraphs.Count - 5));
                            foreach (var para in lastParagraphs)
                            {
                                var paraText = para.InnerText?.Trim();
                                if (!string.IsNullOrWhiteSpace(paraText))
                                {
                                    extractedText.AppendLine($"SIGNATURE: {paraText}");
                                    
                                    if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                                        break;
                                }
                            }
                        }
                        
                        var result = extractedText.ToString();
                        _logger.Debug("OpenXML extraction completed: {Length} characters in {ElapsedMs}ms", 
                            result.Length, stopwatch.ElapsedMilliseconds);
                            
                        return result;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "OpenXML extraction failed for: {Path}", filePath);
                    return string.Empty;
                }
                finally
                {
                    stopwatch.Stop();
                }
            }).ConfigureAwait(false);
        }

        private async Task<string> ExtractTextFromWord(string filePath)
        {
            try
            {
                // Validate file before processing
                var fileInfo = new FileInfo(filePath);
                if (!fileInfo.Exists)
                {
                    _logger.Warning("Word file does not exist: {Path}", filePath);
                    return string.Empty;
                }
                
                if (fileInfo.Length == 0)
                {
                    _logger.Warning("Word file is empty: {Path}", filePath);
                    return string.Empty;
                }
                
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                
                // Try fast OpenXML extraction first for .docx files
                if (extension == ".docx")
                {
                    _logger.Debug("Attempting optimized OpenXML text extraction for: {Path}", filePath);
                    var text = await ExtractTextFromWordDirect(filePath).ConfigureAwait(false);
                    
                    if (!string.IsNullOrWhiteSpace(text) && text.Length > 50)
                    {
                        _logger.Information("Successfully extracted text using optimized OpenXML: {Length} characters", text.Length);
                        return text;
                    }
                    
                    _logger.Information("OpenXML extraction insufficient ({Length} chars), falling back to PDF conversion", text.Length);
                }
                else
                {
                    _logger.Debug("Legacy Word format detected ({Extension}), using PDF conversion", extension);
                }
                
                // Fall back to existing PDF conversion method for .doc files or insufficient OpenXML extraction
                return await ExtractTextFromWordUsingPdfConversion(filePath).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from Word document: {Path}", filePath);
                return string.Empty;
            }
        }

        // Rename the existing ExtractTextFromWord content to this new method
        private async Task<string> ExtractTextFromWordUsingPdfConversion(string filePath)
        {
            try
            {
                // Check if we already have a cached PDF
                if (TryGetCachedPdf(filePath, out var cachedPdfPath) && cachedPdfPath != null)
                {
                    _logger.Information("Using cached PDF for text extraction: {Path}", cachedPdfPath);
                    return await ExtractTextFromPdf(cachedPdfPath).ConfigureAwait(false);
                }
                
                // Validate file before processing
                var fileInfo = new FileInfo(filePath);
                if (!fileInfo.Exists)
                {
                    _logger.Warning("Word file does not exist: {Path}", filePath);
                    return string.Empty;
                }
                
                if (fileInfo.Length == 0)
                {
                    _logger.Warning("Word file is empty: {Path}", filePath);
                    return string.Empty;
                }
                
                // Enterprise security: reasonable file size limit (50MB) to prevent DoS attacks
                const long maxFileSize = 50 * 1024 * 1024; // 50MB
                if (fileInfo.Length > maxFileSize)
                {
                    _logger.Warning("Word file exceeds size limit: {Size} bytes (limit: {MaxSize} bytes)", 
                        fileInfo.Length, maxFileSize);
                    return string.Empty;
                }
                
                _logger.Information("Converting Word to PDF for comprehensive text extraction: {Path}", filePath);
                
                // Create secure temporary directory with unique name
                var tempFolderName = $"DocHandler_{Guid.NewGuid():N}";
                var tempFolder = Path.Combine(Path.GetTempPath(), tempFolderName);
                Directory.CreateDirectory(tempFolder);
                
                var tempPdfPath = Path.Combine(tempFolder, $"{Path.GetFileNameWithoutExtension(filePath)}_scan.pdf");
                
                try
                {
                    // Use existing, trusted Office conversion service
                    var officeService = new OfficeConversionService();
                    var conversionResult = await officeService.ConvertWordToPdf(filePath, tempPdfPath).ConfigureAwait(false);
                    
                    if (!conversionResult.Success)
                    {
                        _logger.Warning("Failed to convert Word to PDF: {Error}", conversionResult.ErrorMessage);
                        
                        // Fallback to basic OpenXML extraction (body paragraphs only)
                        _logger.Information("Falling back to basic text extraction for Word document: {Path}", filePath);
                        return await ExtractTextFromWordBasic(filePath).ConfigureAwait(false);
                    }
                    
                    // Verify the PDF was created successfully
                    if (!File.Exists(tempPdfPath))
                    {
                        _logger.Warning("PDF conversion completed but output file not found: {Path}", tempPdfPath);
                        return await ExtractTextFromWordBasic(filePath).ConfigureAwait(false);
                    }
                    
                    // Additional security check on converted PDF
                    var pdfInfo = new FileInfo(tempPdfPath);
                    if (pdfInfo.Length == 0)
                    {
                        _logger.Warning("Converted PDF is empty: {Path}", tempPdfPath);
                        return await ExtractTextFromWordBasic(filePath).ConfigureAwait(false);
                    }
                    
                    // Cache the converted PDF
                    CachePdfConversion(filePath, tempPdfPath);
                    
                    // Use existing, trusted PDF text extraction (captures headers, footers, images, etc.)
                    var extractedText = await ExtractTextFromPdf(tempPdfPath).ConfigureAwait(false);
                    
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        _logger.Warning("No text extracted from converted Word PDF: {Path}", filePath);
                        return await ExtractTextFromWordBasic(filePath).ConfigureAwait(false);
                    }
                    
                    _logger.Information("Successfully extracted comprehensive text from Word via PDF: {Length} characters from {OriginalFile}", 
                        extractedText.Length, Path.GetFileName(filePath));
                    
                    return extractedText;
                }
                catch (Exception)
                {
                    // Clean up on error
                    try
                    {
                        if (Directory.Exists(tempFolder))
                        {
                            Directory.Delete(tempFolder, true);
                        }
                    }
                    catch { }
                    
                    throw;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from Word document: {Path}", filePath);
                
                // Final fallback to basic extraction
                return await ExtractTextFromWordBasic(filePath).ConfigureAwait(false);
            }
        }

        // Keep the original method as a fallback for when PDF conversion fails
        private async Task<string> ExtractTextFromWordBasic(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    // First, verify this is actually a Word document
                    if (!IsValidWordDocument(filePath))
                    {
                        _logger.Warning("File is not a valid Word document: {Path}", filePath);
                        return string.Empty;
                    }

                    using (var doc = WordprocessingDocument.Open(filePath, false))
                    {
                        var text = new System.Text.StringBuilder();
                        var body = doc.MainDocumentPart?.Document?.Body;
                        
                        if (body == null)
                        {
                            _logger.Warning("Word document has no body content: {Path}", filePath);
                            return string.Empty;
                        }
                        
                        int charCount = 0;
                        int paragraphCount = 0;
                        
                        foreach (var paragraph in body.Elements<Paragraph>())
                        {
                            try
                            {
                                var paraText = paragraph.InnerText;
                                if (!string.IsNullOrWhiteSpace(paraText))
                                {
                                    text.AppendLine(paraText);
                                    charCount += paraText.Length;
                                    paragraphCount++;
                                }
                                
                                // Stop after extracting enough text for company detection
                                if (charCount > 5000)
                                {
                                    _logger.Debug("Sufficient text extracted for company detection from {Count} paragraphs", paragraphCount);
                                    break;
                                }
                            }
                            catch (Exception paraEx)
                            {
                                _logger.Warning(paraEx, "Failed to extract text from paragraph");
                            }
                        }
                        
                        var extractedText = text.ToString();
                        _logger.Debug("Basic text extraction from Word document: {Length} characters from {Count} paragraphs", 
                            extractedText.Length, paragraphCount);
                        return extractedText;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed basic text extraction from Word document: {Path}", filePath);
                    return string.Empty;
                }
            }).ConfigureAwait(false);
        }

        private async Task<string> ExtractTextFromExcelDirect(string filePath)
        {
            return await Task.Run(() =>
            {
                var extractedText = new StringBuilder();
                var stopwatch = Stopwatch.StartNew();
                const int maxChars = 5000;
                const int timeoutMs = 2000;
                
                try
                {
                    // Check file extension - only process .xlsx files with OpenXML
                    var extension = Path.GetExtension(filePath).ToLowerInvariant();
                    if (extension != ".xlsx")
                    {
                        _logger.Debug("Skipping OpenXML extraction for non-xlsx file: {Extension}", extension);
                        return string.Empty;
                    }
                    
                    using (var doc = SpreadsheetDocument.Open(filePath, false))
                    {
                        // 1. Extract document properties first (highest priority)
                        try
                        {
                            var extendedProps = doc.ExtendedFilePropertiesPart;
                            if (extendedProps != null)
                            {
                                var props = extendedProps.Properties;
                                if (props != null)
                                {
                                    var company = props.Company?.Text;
                                    
                                    if (!string.IsNullOrWhiteSpace(company))
                                    {
                                        extractedText.AppendLine($"COMPANY: {company}");
                                        _logger.Debug("Found company in Excel document properties: {Company}", company);
                                    }
                                }
                            }
                            
                            var coreProps = doc.PackageProperties;
                            if (coreProps != null)
                            {
                                var creator = coreProps.Creator;
                                
                                if (!string.IsNullOrWhiteSpace(creator))
                                {
                                    extractedText.AppendLine($"AUTHOR: {creator}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            _logger.Debug(ex, "Failed to extract Excel document properties");
                        }
                        
                        // Early termination check
                        if (extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                        {
                            _logger.Debug("Early termination after Excel document properties");
                            return extractedText.ToString();
                        }
                        
                        var workbookPart = doc.WorkbookPart;
                        if (workbookPart == null)
                        {
                            _logger.Warning("Excel document has no workbook: {Path}", filePath);
                            return extractedText.ToString();
                        }
                        
                        // 2. Extract from first worksheet only
                        var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                        if (worksheetPart == null)
                        {
                            _logger.Warning("Excel document has no worksheets: {Path}", filePath);
                            return extractedText.ToString();
                        }
                        
                        var worksheet = worksheetPart.Worksheet;
                        var sheetData = worksheet.GetFirstChild<SheetData>();
                        if (sheetData == null)
                        {
                            _logger.Debug("Excel worksheet has no data: {Path}", filePath);
                            return extractedText.ToString();
                        }
                        
                        // Get shared string table for text lookup
                        var stringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                        
                        int cellCount = 0;
                        const int maxCells = 100;
                        
                        // 3. Extract text from first 100 cells with content
                        foreach (var row in sheetData.Elements<Row>())
                        {
                            if (cellCount >= maxCells || extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                                break;
                                
                            foreach (var cell in row.Elements<Cell>())
                            {
                                if (cellCount >= maxCells || extractedText.Length >= maxChars || stopwatch.ElapsedMilliseconds > timeoutMs)
                                    break;
                                
                                var cellText = GetCellText(cell, stringTable);
                                if (!string.IsNullOrWhiteSpace(cellText))
                                {
                                    // Skip numbers and formulas - focus on text content
                                    if (!IsNumericOrFormula(cellText))
                                    {
                                        extractedText.AppendLine(cellText);
                                        cellCount++;
                                    }
                                }
                            }
                        }
                        
                        var result = extractedText.ToString();
                        _logger.Debug("Excel OpenXML extraction completed: {Length} characters from {CellCount} cells in {ElapsedMs}ms", 
                            result.Length, cellCount, stopwatch.ElapsedMilliseconds);
                            
                        return result;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Excel OpenXML extraction failed for: {Path}", filePath);
                    return string.Empty;
                }
                finally
                {
                    stopwatch.Stop();
                }
            }).ConfigureAwait(false);
        }
        
        private string GetCellText(Cell cell, SharedStringTable? stringTable)
        {
            if (cell.CellValue == null)
                return string.Empty;
            
            var value = cell.CellValue.Text;
            
            // If it's a shared string, look it up in the string table
            if (cell.DataType?.Value == CellValues.SharedString && stringTable != null)
            {
                if (int.TryParse(value, out int index) && index >= 0 && index < stringTable.Count())
                {
                    var stringItem = stringTable.Elements<SharedStringItem>().ElementAtOrDefault(index);
                    return stringItem?.InnerText ?? string.Empty;
                }
            }
            
            return value ?? string.Empty;
        }
        
        private bool IsNumericOrFormula(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return true;
            
            // Skip if it's a number
            if (double.TryParse(text, out _))
                return true;
            
            // Skip if it looks like a formula result or contains only numbers/symbols
            if (text.All(c => char.IsDigit(c) || ".,%-$()".Contains(c)))
                return true;
            
            return false;
        }

        private async Task<string> ExtractTextFromExcel(string filePath)
        {
            try
            {
                // Validate file before processing
                var fileInfo = new FileInfo(filePath);
                if (!fileInfo.Exists)
                {
                    _logger.Warning("Excel file does not exist: {Path}", filePath);
                    return string.Empty;
                }
                
                if (fileInfo.Length == 0)
                {
                    _logger.Warning("Excel file is empty: {Path}", filePath);
                    return string.Empty;
                }
                
                // Enterprise security: reasonable file size limit (50MB) to prevent DoS attacks
                const long maxFileSize = 50 * 1024 * 1024; // 50MB
                if (fileInfo.Length > maxFileSize)
                {
                    _logger.Warning("Excel file exceeds size limit: {Size} bytes (limit: {MaxSize} bytes)", 
                        fileInfo.Length, maxFileSize);
                    return string.Empty;
                }
                
                var extension = Path.GetExtension(filePath).ToLowerInvariant();
                
                // Try fast OpenXML extraction first for .xlsx files
                if (extension == ".xlsx")
                {
                    _logger.Debug("Attempting optimized OpenXML text extraction for Excel: {Path}", filePath);
                    var text = await ExtractTextFromExcelDirect(filePath).ConfigureAwait(false);
                    
                    if (!string.IsNullOrWhiteSpace(text) && text.Length > 20)
                    {
                        _logger.Information("Successfully extracted text using optimized Excel OpenXML: {Length} characters", text.Length);
                        return text;
                    }
                    
                    _logger.Information("Excel OpenXML extraction insufficient ({Length} chars), falling back to PDF conversion", text.Length);
                }
                else
                {
                    _logger.Debug("Legacy Excel format detected ({Extension}), using PDF conversion", extension);
                }
                
                // Fall back to existing PDF conversion method for .xls files or insufficient OpenXML extraction
                return await ExtractTextFromExcelUsingPdfConversion(filePath).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from Excel document: {Path}", filePath);
                return string.Empty;
            }
        }
        
        // Rename the existing ExtractTextFromExcel content to this new method
        private async Task<string> ExtractTextFromExcelUsingPdfConversion(string filePath)
        {
            try
            {
                _logger.Information("Converting Excel to PDF for text extraction: {Path}", filePath);
                
                // Create secure temporary directory with unique name
                var tempFolderName = $"DocHandler_{Guid.NewGuid():N}";
                var tempFolder = Path.Combine(Path.GetTempPath(), tempFolderName);
                Directory.CreateDirectory(tempFolder);
                
                var tempPdfPath = Path.Combine(tempFolder, "excel_conversion.pdf");
                
                try
                {
                    // Use existing, trusted Office conversion service
                    var officeService = new OfficeConversionService();
                    var conversionResult = await officeService.ConvertExcelToPdf(filePath, tempPdfPath).ConfigureAwait(false);
                    
                    if (!conversionResult.Success)
                    {
                        _logger.Warning("Failed to convert Excel to PDF: {Error}", conversionResult.ErrorMessage);
                        return string.Empty;
                    }
                    
                    // Verify the PDF was created successfully
                    if (!File.Exists(tempPdfPath))
                    {
                        _logger.Warning("PDF conversion completed but output file not found: {Path}", tempPdfPath);
                        return string.Empty;
                    }
                    
                    // Additional security check on converted PDF
                    var pdfInfo = new FileInfo(tempPdfPath);
                    if (pdfInfo.Length == 0)
                    {
                        _logger.Warning("Converted PDF is empty: {Path}", tempPdfPath);
                        return string.Empty;
                    }
                    
                    // Use existing, trusted PDF text extraction
                    var extractedText = await ExtractTextFromPdf(tempPdfPath).ConfigureAwait(false);
                    
                    if (string.IsNullOrWhiteSpace(extractedText))
                    {
                        _logger.Warning("No text extracted from converted Excel PDF: {Path}", filePath);
                        return string.Empty;
                    }
                    
                    _logger.Information("Successfully extracted text from Excel via PDF: {Length} characters from {OriginalFile}", 
                        extractedText.Length, Path.GetFileName(filePath));
                    
                    return extractedText;
                }
                finally
                {
                    // Enterprise-level cleanup: secure removal of temporary files
                    try
                    {
                        if (Directory.Exists(tempFolder))
                        {
                            // Force delete all files in the temporary folder
                            var tempFiles = Directory.GetFiles(tempFolder);
                            foreach (var tempFile in tempFiles)
                            {
                                try
                                {
                                    File.SetAttributes(tempFile, FileAttributes.Normal);
                                    File.Delete(tempFile);
                                }
                                catch (Exception fileEx)
                                {
                                    _logger.Warning(fileEx, "Failed to delete temporary file: {File}", tempFile);
                                }
                            }
                            
                            // Remove the temporary directory
                            Directory.Delete(tempFolder, true);
                            _logger.Debug("Successfully cleaned up temporary folder: {Path}", tempFolder);
                        }
                    }
                    catch (Exception cleanupEx)
                    {
                        _logger.Warning(cleanupEx, "Failed to clean up temporary folder: {Path}", tempFolder);
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from Excel document via PDF conversion: {Path}", filePath);
                return string.Empty;
            }
        }

        private async Task<string> ExtractTextFromTextFile(string filePath)
        {
            try
            {
                // Validate file before attempting to read
                var fileInfo = new FileInfo(filePath);
                if (!fileInfo.Exists)
                {
                    _logger.Warning("Text file does not exist: {Path}", filePath);
                    return string.Empty;
                }
                
                if (fileInfo.Length == 0)
                {
                    _logger.Warning("Text file is empty: {Path}", filePath);
                    return string.Empty;
                }
                
                // Read the text file (limit to first 10KB for company detection)
                var maxBytes = 10 * 1024; // 10KB
                var text = await File.ReadAllTextAsync(filePath).ConfigureAwait(false);
                
                // Truncate if too long for company detection
                if (text.Length > maxBytes)
                {
                    text = text.Substring(0, maxBytes);
                    _logger.Debug("Truncated text file content for company detection: {Length} characters", text.Length);
                }
                
                _logger.Debug("Extracted text from file: {Length} characters", text.Length);
                return text;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from text file: {Path}", filePath);
                return string.Empty;
            }
        }

        private bool IsValidWordDocument(string filePath)
        {
            try
            {
                // Check if it's a valid ZIP file (Office documents are ZIP archives)
                using (var stream = File.OpenRead(filePath))
                {
                    // Check for ZIP signature
                    var signature = new byte[4];
                    if (stream.Read(signature, 0, 4) < 4)
                        return false;
                    
                    // ZIP files start with PK (0x504B)
                    return signature[0] == 0x50 && signature[1] == 0x4B;
                }
            }
            catch
            {
                return false;
            }
        }
        
        public async Task IncrementUsageCount(string companyName)
        {
            var company = _data.Companies.FirstOrDefault(c => 
                c.Name.Equals(companyName, StringComparison.OrdinalIgnoreCase));
            
            if (company != null)
            {
                company.UsageCount++;
                company.LastUsed = DateTime.Now;
                await SaveCompanyNames().ConfigureAwait(false);
            }
        }
        
        public List<CompanyInfo> GetMostUsedCompanies(int count = 10)
        {
            return _data.Companies
                .OrderByDescending(c => c.UsageCount)
                .ThenByDescending(c => c.LastUsed)
                .Take(count)
                .ToList();
        }
        
        public List<CompanyInfo> SearchCompanies(string searchTerm)
        {
            if (string.IsNullOrWhiteSpace(searchTerm))
                return _data.Companies;
            
            searchTerm = searchTerm.ToLowerInvariant();
            
            return _data.Companies
                .Where(c => c.Name.ToLowerInvariant().Contains(searchTerm) ||
                           c.Aliases.Any(a => a.ToLowerInvariant().Contains(searchTerm)))
                .OrderBy(c => c.Name)
                .ToList();
        }
        
        public string GetPerformanceSummary()
        {
            lock (_metricsLock)
            {
                var avgTime = _performanceMetrics.DocumentsProcessed > 0 
                    ? _performanceMetrics.TotalProcessingTime.TotalMilliseconds / _performanceMetrics.DocumentsProcessed 
                    : 0;
                    
                var cacheHitRate = (_performanceMetrics.CacheHits + _performanceMetrics.CacheMisses) > 0
                    ? (double)_performanceMetrics.CacheHits / (_performanceMetrics.CacheHits + _performanceMetrics.CacheMisses) * 100
                    : 0;
                    
                return $"Company Detection Performance: {_performanceMetrics.DocumentsProcessed} docs, " +
                       $"Avg: {avgTime:F1}ms, Cache Hit: {cacheHitRate:F1}%, " +
                       $"PDF Cache: {_convertedPdfCache.Count} entries";
            }
        }

        /// <summary>
        /// Checks if a cached PDF exists for the given file and returns its path
        /// </summary>
        public bool TryGetCachedPdf(string originalFilePath, out string? cachedPdfPath)
        {
            cachedPdfPath = null;
            
            try
            {
                // Check if we have a cached entry
                if (!_convertedPdfCache.TryGetValue(originalFilePath, out var pdfPath))
                {
                    return false;
                }
                
                // Verify the PDF still exists
                if (!File.Exists(pdfPath))
                {
                    _logger.Debug("Cached PDF no longer exists: {Path}", pdfPath);
                    RemoveFromCache(originalFilePath);
                    return false;
                }
                
                // Check if cache is expired
                if (_pdfCacheTimestamps.TryGetValue(originalFilePath, out var timestamp))
                {
                    if (DateTime.Now - timestamp > _cacheExpiration)
                    {
                        _logger.Debug("Cached PDF expired for: {Path}", originalFilePath);
                        RemoveFromCache(originalFilePath);
                        return false;
                    }
                }
                
                // Check if original file has been modified
                if (_pdfCacheFileInfo.TryGetValue(originalFilePath, out var cachedFileInfo))
                {
                    var currentFileInfo = new FileInfo(originalFilePath);
                    if (currentFileInfo.LastWriteTime != cachedFileInfo.LastWriteTime ||
                        currentFileInfo.Length != cachedFileInfo.Length)
                    {
                        _logger.Debug("Original file modified, cache invalidated: {Path}", originalFilePath);
                        RemoveFromCache(originalFilePath);
                        return false;
                    }
                }
                
                cachedPdfPath = pdfPath;
                _logger.Debug("Using cached PDF for: {Path}", originalFilePath);
                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error checking PDF cache for: {Path}", originalFilePath);
                return false;
            }
        }

        /// <summary>
        /// Caches a PDF conversion for later reuse
        /// </summary>
        private void CachePdfConversion(string originalPath, string pdfPath)
        {
            try
            {
                var fileInfo = new FileInfo(originalPath);
                
                _convertedPdfCache[originalPath] = pdfPath;
                _pdfCacheTimestamps[originalPath] = DateTime.Now;
                _pdfCacheFileInfo[originalPath] = fileInfo;
                
                _logger.Debug("Cached PDF conversion: {Original} -> {Pdf}", 
                    Path.GetFileName(originalPath), Path.GetFileName(pdfPath));
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to cache PDF conversion");
            }
        }

        /// <summary>
        /// Removes an entry from the PDF cache
        /// </summary>
        private void RemoveFromCache(string originalPath)
        {
            if (_convertedPdfCache.TryRemove(originalPath, out var pdfPath))
            {
                try
                {
                    if (File.Exists(pdfPath))
                    {
                        File.Delete(pdfPath);
                        _logger.Debug("Deleted cached PDF: {Path}", pdfPath);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Failed to delete cached PDF: {Path}", pdfPath);
                }
            }
            
            _pdfCacheTimestamps.TryRemove(originalPath, out _);
            _pdfCacheFileInfo.TryRemove(originalPath, out _);
        }

        /// <summary>
        /// Cleans up expired PDF cache entries
        /// </summary>
        public void CleanupPdfCache()
        {
            lock (_cacheCleanupLock)
            {
                try
                {
                    _logger.Debug("Starting PDF cache cleanup");
                    var now = DateTime.Now;
                    var expiredCount = 0;
                    var deletedCount = 0;
                    
                    // Find expired entries
                    var expiredKeys = _pdfCacheTimestamps
                        .Where(kvp => now - kvp.Value > _cacheExpiration)
                        .Select(kvp => kvp.Key)
                        .ToList();
                    
                    foreach (var key in expiredKeys)
                    {
                        expiredCount++;
                        if (_convertedPdfCache.TryRemove(key, out var pdfPath))
                        {
                            try
                            {
                                if (File.Exists(pdfPath))
                                {
                                    File.Delete(pdfPath);
                                    deletedCount++;
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.Warning(ex, "Failed to delete expired PDF: {Path}", pdfPath);
                            }
                        }
                        
                        _pdfCacheTimestamps.TryRemove(key, out _);
                        _pdfCacheFileInfo.TryRemove(key, out _);
                    }
                    
                    // Also clean up orphaned PDFs (where original file no longer exists)
                    var orphanedKeys = _convertedPdfCache
                        .Where(kvp => !File.Exists(kvp.Key))
                        .Select(kvp => kvp.Key)
                        .ToList();
                    
                    foreach (var key in orphanedKeys)
                    {
                        RemoveFromCache(key);
                        deletedCount++;
                    }
                    
                    if (expiredCount > 0 || orphanedKeys.Count > 0)
                    {
                        _logger.Information("PDF cache cleanup: {Expired} expired, {Orphaned} orphaned, {Deleted} files deleted", 
                            expiredCount, orphanedKeys.Count, deletedCount);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Error during PDF cache cleanup");
                }
            }
        }

        // Add method to get cached PDF if available
        public string? GetCachedPdfPath(string originalFilePath)
        {
            if (_convertedPdfCache.TryGetValue(originalFilePath, out string? pdfPath) && 
                File.Exists(pdfPath))
            {
                // Check if original file hasn't changed
                var currentInfo = new FileInfo(originalFilePath);
                if (_pdfCacheFileInfo.TryGetValue(originalFilePath, out var cachedInfo) &&
                    currentInfo.LastWriteTime == cachedInfo.LastWriteTime &&
                    currentInfo.Length == cachedInfo.Length)
                {
                    return pdfPath;
                }
                else
                {
                    // Original file changed, invalidate cache
                    _convertedPdfCache.TryRemove(originalFilePath, out _);
                    if (File.Exists(pdfPath))
                    {
                        try { File.Delete(pdfPath); } catch { }
                    }
                }
            }
            
            return null;
        }

        // Add method to remove cached PDF
        public void RemoveCachedPdf(string originalFilePath)
        {
            if (string.IsNullOrWhiteSpace(originalFilePath))
                return;
                
            lock (_cacheCleanupLock)
            {
                if (_convertedPdfCache.TryRemove(originalFilePath, out var cachedPdfPath))
                {
                    _pdfCacheTimestamps.TryRemove(originalFilePath, out _);
                    _pdfCacheFileInfo.TryRemove(originalFilePath, out _);
                    
                    if (File.Exists(cachedPdfPath))
                    {
                        try
                        {
                            File.Delete(cachedPdfPath);
                            _logger.Debug("Removed cached PDF: {Path}", cachedPdfPath);
                        }
                        catch (Exception ex)
                        {
                            _logger.Warning(ex, "Failed to delete cached PDF: {Path}", cachedPdfPath);
                        }
                    }
                }
            }
        }

        private async Task<string?> ProcessDocFileForCompanyDetection(string filePath, IProgress<int>? progress, CancellationToken cancellationToken)
        {
            var fileInfo = new FileInfo(filePath);
            var fileSizeMB = fileInfo.Length / (1024.0 * 1024.0);
            
            // Use configurable size limit for .doc files
            if (fileSizeMB > _docFileSizeLimitMB)
            {
                _logger.Warning("Skipping large .doc file ({Size:F1}MB) exceeding limit of {LimitMB}MB: {Path}", 
                    fileSizeMB, _docFileSizeLimitMB, filePath);
                return null;
            }
            
            progress?.Report(20);
            
            // Try basic text extraction first (faster, safer)
            try
            {
                var basicText = await ExtractTextFromDocBasic(filePath);
                if (!string.IsNullOrWhiteSpace(basicText) && basicText.Length > 100)
                {
                    progress?.Report(70);
                    
                    var processedText = PreprocessText(basicText);
                    var match = await FindBestCompanyMatch(processedText);
                    
                    progress?.Report(100);
                    
                    if (match.HasValue)
                    {
                        _logger.Information("Found company in .doc file using basic extraction: {Company}", match.Value.company);
                        return match.Value.company;
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Basic text extraction failed for .doc file: {Path}", filePath);
            }
            
            // If basic extraction didn't work, try PDF conversion with strict timeout
            progress?.Report(40);
            
            try
            {
                var pdfText = await ExtractTextFromDocViaPdfConversion(filePath, cancellationToken);
                if (!string.IsNullOrWhiteSpace(pdfText))
                {
                    progress?.Report(80);
                    
                    var processedText = PreprocessText(pdfText);
                    var match = await FindBestCompanyMatch(processedText);
                    
                    progress?.Report(100);
                    
                    if (match.HasValue)
                    {
                        _logger.Information("Found company in .doc file using PDF conversion: {Company}", match.Value.company);
                        return match.Value.company;
                    }
                }
            }
            catch (OperationCanceledException)
            {
                _logger.Warning("PDF conversion timed out for .doc file: {Path}", filePath);
                throw;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "PDF conversion failed for .doc file: {Path}", filePath);
            }
            
            progress?.Report(100);
            return null;
        }

        private async Task<string> ExtractTextFromDocBasic(string filePath)
        {
            // Try to extract text using basic file reading (for simple .doc files)
            try
            {
                var bytes = await File.ReadAllBytesAsync(filePath);
                var text = System.Text.Encoding.UTF8.GetString(bytes);
                
                // Basic cleanup of binary data
                text = Regex.Replace(text, @"[^\x20-\x7E\r\n]", " ");
                text = Regex.Replace(text, @"\s+", " ");
                
                if (text.Length > 1000)
                {
                    text = text.Substring(0, 1000); // Limit to first 1000 characters
                }
                
                return text;
            }
            catch (Exception ex)
            {
                _logger.Debug(ex, "Basic text extraction failed for .doc file");
                return string.Empty;
            }
        }

        private async Task<string> ExtractTextFromDocViaPdfConversion(string filePath, CancellationToken cancellationToken)
        {
            if (_sessionOfficeService == null)
            {
                _logger.Warning("Office service not available for .doc PDF conversion");
                return string.Empty;
            }
            
            var tempPdf = Path.Combine(Path.GetTempPath(), $"DocHandler_{Guid.NewGuid()}.pdf");
            
            try
            {
                // Use a task with timeout for the conversion
                var conversionTask = _sessionOfficeService.ConvertWordToPdf(filePath, tempPdf);
                var timeoutTask = Task.Delay(TimeSpan.FromSeconds(10), cancellationToken);
                
                var completedTask = await Task.WhenAny(conversionTask, timeoutTask);
                
                if (completedTask == timeoutTask)
                {
                    _logger.Warning("PDF conversion timed out for .doc file: {Path}", filePath);
                    throw new OperationCanceledException("PDF conversion timed out");
                }
                
                var conversionResult = await conversionTask;
                
                if (!conversionResult.Success)
                {
                    _logger.Warning("Failed to convert .doc to PDF: {Error}", conversionResult.ErrorMessage);
                    return string.Empty;
                }
                
                if (!File.Exists(tempPdf))
                {
                    _logger.Warning("PDF conversion completed but output file not found: {Path}", tempPdf);
                    return string.Empty;
                }
                
                // Extract text from the converted PDF
                return await ExtractTextFromPdf(tempPdf);
            }
            finally
            {
                // Clean up temporary PDF
                if (File.Exists(tempPdf))
                {
                    try
                    {
                        File.Delete(tempPdf);
                    }
                    catch (Exception ex)
                    {
                        _logger.Warning(ex, "Failed to delete temporary PDF: {Path}", tempPdf);
                    }
                }
            }
        }
        
        public void Dispose()
        {
            try
            {
                _cacheCleanupTimer?.Dispose();
                
                // Clean up all cached PDFs on disposal
                _logger.Information("Cleaning up PDF cache on disposal");
                
                foreach (var kvp in _convertedPdfCache)
                {
                    try
                    {
                        if (File.Exists(kvp.Value))
                        {
                            File.Delete(kvp.Value);
                        }
                    }
                    catch { }
                }
                
                _convertedPdfCache.Clear();
                _pdfCacheTimestamps.Clear();
                _pdfCacheFileInfo.Clear();
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error during CompanyNameService disposal");
            }
        }
    }
    
    public class CompanyNamesData
    {
        public List<CompanyInfo> Companies { get; set; } = new();
    }
    
    public class CompanyInfo
    {
        public string Name { get; set; } = "";
        public List<string> Aliases { get; set; } = new();
        public DateTime DateAdded { get; set; } = DateTime.Now;
        public DateTime? LastUsed { get; set; }
        public int UsageCount { get; set; }
    }
}