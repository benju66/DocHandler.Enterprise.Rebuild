using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Serilog;
using Task = System.Threading.Tasks.Task;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocHandler.Services
{
    public class CompanyNameService
    {
        private readonly ILogger _logger;
        private readonly string _dataPath;
        private readonly string _companyNamesPath;
        private CompanyNamesData _data;
        
        public List<CompanyInfo> Companies => _data.Companies;
        
        public CompanyNameService()
        {
            _logger = Log.ForContext<CompanyNameService>();
            
            // Store data in AppData
            var appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            var appFolder = Path.Combine(appDataPath, "DocHandler");
            Directory.CreateDirectory(appFolder);
            
            _dataPath = appFolder;
            _companyNamesPath = Path.Combine(appFolder, "company_names.json");
            _data = LoadCompanyNames();
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
                await File.WriteAllTextAsync(_companyNamesPath, json);
                
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
                
                await SaveCompanyNames();
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
                    await SaveCompanyNames();
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
        
        public async Task<string?> ScanDocumentForCompanyName(string filePath)
        {
            try
            {
                _logger.Information("Scanning document for company names: {Path}", filePath);
                
                // Validate file exists and is accessible
                if (!File.Exists(filePath))
                {
                    _logger.Warning("File does not exist: {Path}", filePath);
                    return null;
                }
                
                // Extract text from the document
                string documentText = await ExtractTextFromDocument(filePath);
                
                if (string.IsNullOrWhiteSpace(documentText))
                {
                    _logger.Warning("No text extracted from document: {Path}", filePath);
                    return null;
                }
                
                // Validate that we have meaningful text (more than just whitespace/special chars)
                if (documentText.Trim().Length < 10)
                {
                    _logger.Warning("Insufficient text content for company detection: {Length} characters", documentText.Trim().Length);
                    return null;
                }
                
                _logger.Debug("Extracted {Length} characters from document for company detection", documentText.Length);
                
                // Convert to lowercase for comparison
                string lowerText = documentText.ToLowerInvariant();
                
                // Check if we have any companies to search for
                if (!_data.Companies.Any())
                {
                    _logger.Warning("No companies in database to search for");
                    return null;
                }
                
                // Check each company and its aliases, ordered by usage count for efficiency
                foreach (var company in _data.Companies.OrderByDescending(c => c.UsageCount))
                {
                    // Skip companies with empty names
                    if (string.IsNullOrWhiteSpace(company.Name))
                        continue;
                    
                    // Check main name
                    if (ContainsCompanyName(lowerText, company.Name))
                    {
                        _logger.Information("Found company name: {Name}", company.Name);
                        await IncrementUsageCount(company.Name);
                        return company.Name;
                    }
                    
                    // Check aliases
                    foreach (var alias in company.Aliases)
                    {
                        if (!string.IsNullOrWhiteSpace(alias) && ContainsCompanyName(lowerText, alias))
                        {
                            _logger.Information("Found company via alias '{Alias}': {Name}", alias, company.Name);
                            await IncrementUsageCount(company.Name);
                            return company.Name;
                        }
                    }
                }
                
                _logger.Information("No known company names found in document (searched {Count} companies)", _data.Companies.Count);
                return null;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to scan document for company names: {Path}", filePath);
                return null;
            }
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
        
        private async Task<string> ExtractTextFromDocument(string filePath)
        {
            var extension = Path.GetExtension(filePath).ToLowerInvariant();
            
            try
            {
                switch (extension)
                {
                    case ".pdf":
                        return await ExtractTextFromPdf(filePath);
                    case ".doc":
                    case ".docx":
                        return await ExtractTextFromWord(filePath);
                    case ".txt":
                        return await ExtractTextFromTextFile(filePath);
                    default:
                        _logger.Warning("Unsupported file type for text extraction: {Extension}", extension);
                        return string.Empty;
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to extract text from document: {Path}", filePath);
                return string.Empty;
            }
        }
        
        private async Task<string> ExtractTextFromPdf(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    // Validate file before attempting to read
                    var fileInfo = new FileInfo(filePath);
                    if (!fileInfo.Exists)
                    {
                        _logger.Warning("PDF file does not exist: {Path}", filePath);
                        return string.Empty;
                    }
                    
                    if (fileInfo.Length == 0)
                    {
                        _logger.Warning("PDF file is empty: {Path}", filePath);
                        return string.Empty;
                    }
                    
                    using (var reader = new PdfReader(filePath))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        var text = new System.Text.StringBuilder();
                        
                        // Only extract first 3 pages for company detection (optimization)
                        int totalPages = pdfDoc.GetNumberOfPages();
                        int maxPages = Math.Min(3, totalPages);
                        
                        _logger.Debug("Extracting text from PDF: {Pages} pages (max {Max})", totalPages, maxPages);
                        
                        for (int page = 1; page <= maxPages; page++)
                        {
                            try
                            {
                                var strategy = new SimpleTextExtractionStrategy();
                                var pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                                
                                if (!string.IsNullOrWhiteSpace(pageText))
                                {
                                    text.AppendLine(pageText);
                                    _logger.Debug("Extracted {Length} characters from page {Page}", pageText.Length, page);
                                }
                                
                                // If we've already found enough text (e.g., 5000 characters), stop scanning
                                if (text.Length > 5000)
                                {
                                    _logger.Debug("Sufficient text extracted for company detection, stopping at page {Page}", page);
                                    break;
                                }
                            }
                            catch (Exception pageEx)
                            {
                                _logger.Warning(pageEx, "Failed to extract text from page {Page} of PDF", page);
                            }
                        }
                        
                        var extractedText = text.ToString();
                        _logger.Debug("Total text extracted from PDF: {Length} characters", extractedText.Length);
                        return extractedText;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to extract text from PDF: {Path}", filePath);
                    return string.Empty;
                }
            });
        }
        
        private async Task<string> ExtractTextFromWord(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    // Validate file before attempting to read
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
                        _logger.Debug("Total text extracted from Word document: {Length} characters from {Count} paragraphs", extractedText.Length, paragraphCount);
                        return extractedText;
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to extract text from Word document: {Path}", filePath);
                    return string.Empty;
                }
            });
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
                var text = await File.ReadAllTextAsync(filePath);
                
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
                await SaveCompanyNames();
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