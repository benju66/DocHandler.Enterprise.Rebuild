using System;
using System.IO;
using System.Linq;
using DocHandler.Services;
using Serilog;

namespace DocHandler
{
    /// <summary>
    /// Console application for running memory leak analysis
    /// </summary>
    public class MemoryLeakAnalyzer
    {
        private static void Main(string[] args)
        {
            // Configure logging
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.Console()
                .WriteTo.File("memory-leak-analysis.log")
                .CreateLogger();

            try
            {
                Console.WriteLine("🔍 DocHandler Memory Leak Analyzer");
                Console.WriteLine("==================================\n");

                var detector = new MemoryLeakDetector();
                string targetPath = args.Length > 0 ? args[0] : Directory.GetCurrentDirectory();

                if (!Directory.Exists(targetPath) && !File.Exists(targetPath))
                {
                    Console.WriteLine($"❌ Path not found: {targetPath}");
                    return;
                }

                Console.WriteLine($"📁 Analyzing path: {targetPath}");
                Console.WriteLine("⏳ Running analysis...\n");

                var results = File.Exists(targetPath) 
                    ? detector.AnalyzeFile(targetPath).ToList()
                    : detector.AnalyzeDirectory(targetPath);

                var report = detector.GenerateReport(results);
                Console.WriteLine(report);

                // Save detailed results if any issues found
                if (results.Any())
                {
                    var detailedReport = GenerateDetailedReport(results);
                    var reportPath = Path.Combine(
                        Directory.GetCurrentDirectory(), 
                        $"memory-leak-detailed-{DateTime.Now:yyyyMMdd-HHmmss}.txt"
                    );
                    File.WriteAllText(reportPath, detailedReport);
                    Console.WriteLine($"\n📄 Detailed report saved to: {reportPath}");
                }

                Console.WriteLine("\n✅ Analysis complete!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error during analysis: {ex.Message}");
                Log.Error(ex, "Error during memory leak analysis");
            }
            finally
            {
                Log.CloseAndFlush();
            }
        }

        private static string GenerateDetailedReport(System.Collections.Generic.List<MemoryLeakDetector.DetectionResult> results)
        {
            var report = new System.Text.StringBuilder();
            report.AppendLine("DETAILED MEMORY LEAK ANALYSIS REPORT");
            report.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            report.AppendLine("=" + new string('=', 50));
            report.AppendLine();

            foreach (var group in results.GroupBy(r => r.FilePath))
            {
                report.AppendLine($"FILE: {group.Key}");
                report.AppendLine("-" + new string('-', group.Key.Length + 5));
                
                foreach (var result in group.OrderBy(r => r.LineNumber))
                {
                    report.AppendLine($"  🔸 Line {result.LineNumber} [{result.Severity}]");
                    report.AppendLine($"     Message: {result.Message}");
                    report.AppendLine($"     Pattern: {result.RulePattern}");
                    report.AppendLine($"     Code: {result.LineContent}");
                    report.AppendLine();
                }
                report.AppendLine();
            }

            return report.ToString();
        }
    }
}