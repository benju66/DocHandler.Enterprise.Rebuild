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
                
                // Extract text from the document
                string documentText = await ExtractTextFromDocument(filePath);
                
                if (string.IsNullOrWhiteSpace(documentText))
                {
                    _logger.Warning("No text extracted from document");
                    return null;
                }
                
                // Convert to lowercase for comparison
                string lowerText = documentText.ToLowerInvariant();
                
                // Check each company and its aliases, ordered by usage count for efficiency
                foreach (var company in _data.Companies.OrderByDescending(c => c.UsageCount))
                {
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
                
                _logger.Information("No known company names found in document");
                return null;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to scan document for company names");
                return null;
            }
        }
        
        private bool ContainsCompanyName(string text, string companyName)
        {
            try
            {
                // Simple word boundary match - can be enhanced with better patterns
                string pattern = $@"\b{Regex.Escape(companyName.ToLowerInvariant())}\b";
                return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Regex match failed for company name: {Name}", companyName);
                // Fallback to simple contains
                return text.Contains(companyName.ToLowerInvariant());
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
                    using (var reader = new PdfReader(filePath))
                    using (var pdfDoc = new PdfDocument(reader))
                    {
                        var text = new System.Text.StringBuilder();
                        
                        // Only extract first 3 pages for company detection (optimization)
                        int maxPages = Math.Min(3, pdfDoc.GetNumberOfPages());
                        
                        for (int page = 1; page <= maxPages; page++)
                        {
                            try
                            {
                                var strategy = new SimpleTextExtractionStrategy();
                                var pageText = PdfTextExtractor.GetTextFromPage(pdfDoc.GetPage(page), strategy);
                                text.AppendLine(pageText);
                                
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
                        
                        return text.ToString();
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to extract text from PDF");
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
                        
                        if (body != null)
                        {
                            int charCount = 0;
                            foreach (var paragraph in body.Elements<Paragraph>())
                            {
                                try
                                {
                                    var paraText = paragraph.InnerText;
                                    text.AppendLine(paraText);
                                    charCount += paraText.Length;
                                    
                                    // Stop after extracting enough text for company detection
                                    if (charCount > 5000)
                                    {
                                        _logger.Debug("Sufficient text extracted for company detection");
                                        break;
                                    }
                                }
                                catch (Exception paraEx)
                                {
                                    _logger.Warning(paraEx, "Failed to extract text from paragraph");
                                }
                            }
                        }
                        
                        return text.ToString();
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Failed to extract text from Word document");
                    return string.Empty;
                }
            });
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