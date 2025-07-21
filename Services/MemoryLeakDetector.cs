using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Detects potential memory leaks and disposal issues in C# code
    /// </summary>
    public class MemoryLeakDetector
    {
        private readonly ILogger _logger;
        private readonly List<DetectionRule> _rules;
        private readonly bool _isEnabled;
        private readonly DetectionMode _mode;

        public MemoryLeakDetector()
        {
            _logger = Log.ForContext<MemoryLeakDetector>();
            _isEnabled = true;
            _mode = DetectionMode.DetectOnly;
            _rules = InitializeRules();
        }

        public enum DetectionMode
        {
            DetectOnly,
            DetectAndFix
        }

        public enum Severity
        {
            Warning,
            Critical
        }

        public class DetectionRule
        {
            public string Pattern { get; set; }
            public string Message { get; set; }
            public Severity Severity { get; set; }
            public Regex CompiledPattern { get; set; }
            
            public DetectionRule(string pattern, string message, Severity severity)
            {
                Pattern = pattern;
                Message = message;
                Severity = severity;
                CompiledPattern = new Regex(pattern, RegexOptions.Compiled | RegexOptions.Multiline);
            }
        }

        public class DetectionResult
        {
            public string FilePath { get; set; }
            public int LineNumber { get; set; }
            public string LineContent { get; set; }
            public string Message { get; set; }
            public Severity Severity { get; set; }
            public string RulePattern { get; set; }
        }

        private List<DetectionRule> InitializeRules()
        {
            return new List<DetectionRule>
            {
                // Object creation without using statement
                new DetectionRule(
                    @"new\s+\w+\(\)(?!.*using\s*\()",
                    "Object creation without using statement",
                    Severity.Warning
                ),
                
                // COM collection access - verify disposal
                new DetectionRule(
                    @"_\w+App\.\w+\.Open\([^)]*\)",
                    "COM collection access - verify disposal",
                    Severity.Critical
                ),
                
                // Office collection access - track disposal
                new DetectionRule(
                    @"\.Documents\.|\.Workbooks\.",
                    "Office collection access - track disposal",
                    Severity.Critical
                ),
                
                // Event subscription without unsubscribe
                new DetectionRule(
                    @"\+=(?!.*\-=)",
                    "Event subscription without unsubscribe",
                    Severity.Warning
                )
            };
        }

        /// <summary>
        /// Analyzes a single file for memory leak patterns
        /// </summary>
        public List<DetectionResult> AnalyzeFile(string filePath)
        {
            if (!_isEnabled || !File.Exists(filePath))
                return new List<DetectionResult>();

            var results = new List<DetectionResult>();
            
            try
            {
                var lines = File.ReadAllLines(filePath);
                var fileContent = File.ReadAllText(filePath);
                
                foreach (var rule in _rules)
                {
                    var matches = rule.CompiledPattern.Matches(fileContent);
                    
                    foreach (Match match in matches)
                    {
                        var lineNumber = GetLineNumber(fileContent, match.Index);
                        var lineContent = lineNumber <= lines.Length ? lines[lineNumber - 1].Trim() : "";
                        
                        // Skip if this is a known safe pattern
                        if (IsSafePattern(lineContent, rule, filePath))
                            continue;
                            
                        results.Add(new DetectionResult
                        {
                            FilePath = filePath,
                            LineNumber = lineNumber,
                            LineContent = lineContent,
                            Message = rule.Message,
                            Severity = rule.Severity,
                            RulePattern = rule.Pattern
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error analyzing file {FilePath}", filePath);
            }
            
            return results;
        }

        /// <summary>
        /// Analyzes all C# files in a directory
        /// </summary>
        public List<DetectionResult> AnalyzeDirectory(string directoryPath, bool recursive = true)
        {
            if (!_isEnabled || !Directory.Exists(directoryPath))
                return new List<DetectionResult>();

            var results = new List<DetectionResult>();
            var searchOption = recursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            
            try
            {
                var csharpFiles = Directory.GetFiles(directoryPath, "*.cs", searchOption)
                    .Where(f => !f.Contains("\\bin\\") && !f.Contains("\\obj\\"))
                    .ToList();

                _logger.Information("Analyzing {FileCount} C# files in {Directory}", csharpFiles.Count, directoryPath);

                foreach (var file in csharpFiles)
                {
                    var fileResults = AnalyzeFile(file);
                    results.AddRange(fileResults);
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error analyzing directory {DirectoryPath}", directoryPath);
            }
            
            return results;
        }

        /// <summary>
        /// Checks if a detected pattern is actually safe and should be ignored
        /// </summary>
        private bool IsSafePattern(string lineContent, DetectionRule rule, string filePath)
        {
            // Skip if line is commented out
            if (lineContent.TrimStart().StartsWith("//"))
                return true;

            // Safe patterns for COM collection access
            if (rule.Pattern.Contains("Documents|Workbooks"))
            {
                // Allow if using ComResourceScope
                if (lineContent.Contains("comScope.Track") || 
                    lineContent.Contains("ComResourceScope") ||
                    lineContent.Contains("using (var comScope"))
                    return true;
                    
                // Allow if it's in a method that properly disposes
                if (lineContent.Contains("GetDocuments") || lineContent.Contains("GetWorkbooks"))
                    return true;
            }

            // Safe patterns for object creation
            if (rule.Pattern.Contains("new\\s+\\w+\\(\\)"))
            {
                // Allow value types, strings, and simple objects
                if (Regex.IsMatch(lineContent, @"new\s+(int|string|bool|DateTime|TimeSpan|StringBuilder|List<|Dictionary<|ConcurrentQueue<)"))
                    return true;
                    
                // Allow if it's assigned to a using statement
                if (lineContent.Contains("using var") || lineContent.Contains("using ("))
                    return true;
                    
                // Allow if it's in a using block context
                if (IsInUsingBlock(filePath, lineContent))
                    return true;
            }

            // Safe patterns for event subscriptions
            if (rule.Pattern.Contains("\\+="))
            {
                // Allow if there's a corresponding unsubscribe in the same class/method
                if (HasCorrespondingUnsubscribe(filePath, lineContent))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Checks if the line is within a using block
        /// </summary>
        private bool IsInUsingBlock(string filePath, string lineContent)
        {
            // This is a simplified check - in a real implementation you'd parse the syntax tree
            try
            {
                var fileContent = File.ReadAllText(filePath);
                var lines = fileContent.Split('\n');
                
                // Look for using blocks around this line
                foreach (var line in lines)
                {
                    if (line.Contains(lineContent) && 
                        fileContent.Contains("using (") && 
                        fileContent.Contains("Dispose"))
                        return true;
                }
            }
            catch { }
            
            return false;
        }

        /// <summary>
        /// Checks if there's a corresponding event unsubscribe
        /// </summary>
        private bool HasCorrespondingUnsubscribe(string filePath, string lineContent)
        {
            try
            {
                var fileContent = File.ReadAllText(filePath);
                
                // Extract event name from subscription
                var match = Regex.Match(lineContent, @"(\w+)\s*\+=");
                if (match.Success)
                {
                    var eventName = match.Groups[1].Value;
                    return fileContent.Contains($"{eventName} -=");
                }
            }
            catch { }
            
            return false;
        }

        /// <summary>
        /// Gets the line number for a character position in text
        /// </summary>
        private int GetLineNumber(string text, int position)
        {
            return text.Take(position).Count(c => c == '\n') + 1;
        }

        /// <summary>
        /// Generates a detailed report of all detection results
        /// </summary>
        public string GenerateReport(List<DetectionResult> results)
        {
            if (!results.Any())
                return "✅ No memory leak issues detected.";

            var report = new System.Text.StringBuilder();
            report.AppendLine("🔍 Memory Leak Detection Report");
            report.AppendLine("=" + new string('=', 40));
            report.AppendLine();

            var criticalIssues = results.Where(r => r.Severity == Severity.Critical).ToList();
            var warningIssues = results.Where(r => r.Severity == Severity.Warning).ToList();

            report.AppendLine($"📊 Summary:");
            report.AppendLine($"   Critical Issues: {criticalIssues.Count}");
            report.AppendLine($"   Warnings: {warningIssues.Count}");
            report.AppendLine($"   Total Issues: {results.Count}");
            report.AppendLine();

            if (criticalIssues.Any())
            {
                report.AppendLine("🚨 CRITICAL ISSUES:");
                report.AppendLine("-" + new string('-', 30));
                foreach (var issue in criticalIssues.GroupBy(i => i.FilePath))
                {
                    report.AppendLine($"\n📁 {Path.GetFileName(issue.Key)}:");
                    foreach (var result in issue.OrderBy(i => i.LineNumber))
                    {
                        report.AppendLine($"   Line {result.LineNumber}: {result.Message}");
                        report.AppendLine($"   Code: {result.LineContent}");
                        report.AppendLine();
                    }
                }
            }

            if (warningIssues.Any())
            {
                report.AppendLine("⚠️  WARNINGS:");
                report.AppendLine("-" + new string('-', 20));
                foreach (var issue in warningIssues.GroupBy(i => i.FilePath))
                {
                    report.AppendLine($"\n📁 {Path.GetFileName(issue.Key)}:");
                    foreach (var result in issue.OrderBy(i => i.LineNumber))
                    {
                        report.AppendLine($"   Line {result.LineNumber}: {result.Message}");
                        report.AppendLine($"   Code: {result.LineContent}");
                        report.AppendLine();
                    }
                }
            }

            report.AppendLine("💡 Recommendations:");
            report.AppendLine("   - Use ComResourceScope for COM object management");
            report.AppendLine("   - Wrap disposable objects in using statements");
            report.AppendLine("   - Ensure event handlers are unsubscribed");
            report.AppendLine("   - Consider implementing IDisposable for classes with unmanaged resources");

            return report.ToString();
        }

        /// <summary>
        /// Analyzes the current project for memory leak issues
        /// </summary>
        public void AnalyzeProject()
        {
            if (!_isEnabled)
            {
                _logger.Information("Memory leak detector is disabled");
                return;
            }

            _logger.Information("Starting memory leak analysis...");
            
            var projectRoot = Directory.GetCurrentDirectory();
            var results = AnalyzeDirectory(projectRoot);
            
            var report = GenerateReport(results);
            _logger.Information("Memory leak analysis completed. Results:\n{Report}", report);
            
            // Optionally save report to file
            var reportPath = Path.Combine(projectRoot, "memory-leak-report.txt");
            File.WriteAllText(reportPath, report);
            _logger.Information("Report saved to {ReportPath}", reportPath);
        }
    }
}