using System;
using System.Threading.Tasks;
using System.IO;
using DocHandler.Services;
using DocHandler.Models;
using Serilog;
using System.Threading;
using System.Windows;

namespace DocHandler
{
    public static class QuickDiagnostic
    {
        private static readonly ILogger _logger = Log.ForContext(typeof(QuickDiagnostic));
        
        public static async Task<string> TestQueueProcessing()
        {
            var results = new System.Text.StringBuilder();
            results.AppendLine("=== QUEUE PROCESSING DIAGNOSTIC ===");
            
            try
            {
                // Test 1: Service Dependencies
                results.AppendLine("\n🔧 Testing Service Dependencies...");
                var (servicesOk, servicesError) = await TestServiceDependencies();
                results.AppendLine(servicesOk ? "✅ Services: OK" : $"❌ Services: {servicesError}");
                
                if (!servicesOk)
                {
                    results.AppendLine("\n🚨 CRITICAL: Service initialization failed. Queue cannot work.");
                    return results.ToString();
                }
                
                // Test 2: Office Availability
                results.AppendLine("\n🏢 Testing Office Availability...");
                var (officeOk, officeError) = await TestOfficeAvailability();
                results.AppendLine(officeOk ? "✅ Office: Available" : $"❌ Office: {officeError}");
                
                if (!officeOk)
                {
                    results.AppendLine("\n🚨 CRITICAL: Microsoft Office not available. Queue cannot process Office files.");
                    return results.ToString();
                }
                
                // Test 3: Queue Service Creation
                results.AppendLine("\n📋 Testing Queue Service Creation...");
                var (queueOk, queueError) = await TestQueueService();
                results.AppendLine(queueOk ? "✅ Queue Service: OK" : $"❌ Queue Service: {queueError}");
                
                if (!queueOk)
                {
                    results.AppendLine("\n🚨 CRITICAL: Queue service creation failed.");
                    return results.ToString();
                }
                
                // Test 4: STA Threading
                results.AppendLine("\n🧵 Testing STA Threading...");
                var staOk = await TestStaThreading();
                results.AppendLine(staOk ? "✅ STA Threading: OK" : "❌ STA Threading: Failed");
                
                if (!staOk)
                {
                    results.AppendLine("\n🚨 CRITICAL: STA threading failed. COM operations will fail.");
                    return results.ToString();
                }
                
                // Test 5: Simple Queue Operation
                results.AppendLine("\n🔄 Testing Queue Operations...");
                var (operationOk, operationError) = await TestQueueOperations();
                results.AppendLine(operationOk ? "✅ Queue Operations: OK" : $"❌ Queue Operations: {operationError}");
                
                if (operationOk)
                {
                    results.AppendLine("\n✅ ALL TESTS PASSED - Queue should be working!");
                    results.AppendLine("\nIf files still don't process, check:");
                    results.AppendLine("- File permissions on input/output folders");
                    results.AppendLine("- Application logs for specific error messages");
                    results.AppendLine("- Try restarting the application");
                }
                else
                {
                    results.AppendLine($"\n🚨 Queue operations failed: {operationError}");
                }
                
                return results.ToString();
            }
            catch (Exception ex)
            {
                results.AppendLine($"\n💥 FATAL ERROR: {ex.Message}");
                _logger.Error(ex, "Fatal error during diagnostic");
                return results.ToString();
            }
        }
        
        private static async Task<(bool success, string error)> TestServiceDependencies()
        {
            try
            {
                var configService = new ConfigurationService();
                var processManager = new ProcessManager();
                var pdfCacheService = new PdfCacheService();
                var officeTracker = new OfficeInstanceTracker();
                
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
        
        private static async Task<(bool success, string error)> TestOfficeAvailability()
        {
            try
            {
                // Test Word
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    return (false, "Word.Application ProgID not found - Microsoft Word is not installed or registered");
                }
                
                // Test Excel
                Type? excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    return (false, "Excel.Application ProgID not found - Microsoft Excel is not installed or registered");
                }
                
                // Test actual Word creation
                bool wordWorks = false;
                Exception? wordException = null;
                
                var thread = new Thread(() =>
                {
                    try
                    {
                        dynamic wordApp = Activator.CreateInstance(wordType);
                        wordApp.Visible = false;
                        wordApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                        wordWorks = true;
                    }
                    catch (Exception ex)
                    {
                        wordException = ex;
                    }
                });
                
                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
                
                if (!wordWorks)
                {
                    return (false, $"Word application creation failed: {wordException?.Message ?? "Unknown error"}");
                }
                
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
        
        private static async Task<(bool success, string error)> TestQueueService()
        {
            try
            {
                var configService = new ConfigurationService();
                var pdfCacheService = new PdfCacheService();
                var processManager = new ProcessManager();
                var officeTracker = new OfficeInstanceTracker();
                
                var queueService = new SaveQuotesQueueService(configService, pdfCacheService, processManager, officeTracker);
                
                // Test basic properties
                bool isProcessing = queueService.IsProcessing;
                int totalCount = queueService.TotalCount;
                
                queueService.Dispose();
                
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
        
        private static async Task<bool> TestStaThreading()
        {
            try
            {
                Log.Information("Testing STA Threading...");
                
                // Test 1: Create a simple STA thread pool
                using var testPool = new StaThreadPool(1, "DiagnosticTest");
                
                // Wait a moment for thread initialization
                await Task.Delay(500); // Increased delay to ensure thread startup
                
                // Test 2: Verify threads are actually STA and alive
                if (!testPool.VerifyStaThreads())
                {
                    Log.Error("STA Threading: STA thread pool verification failed - threads are not STA or not alive");
                    return false;
                }
                
                // Test 3: Test actual functionality
                var functionalityTest = await testPool.TestFunctionality();
                if (!functionalityTest)
                {
                    Log.Error("STA Threading: Thread pool functionality test failed");
                    return false;
                }
                
                // Test 4: Execute a simple operation on STA thread
                bool operationSucceeded = false;
                try
                {
                    await testPool.ExecuteAsync(() =>
                    {
                        // Verify we're on STA thread
                        var apartmentState = Thread.CurrentThread.GetApartmentState();
                        if (apartmentState != ApartmentState.STA)
                        {
                            throw new InvalidOperationException($"Expected STA thread, got {apartmentState}");
                        }
                        
                        operationSucceeded = true;
                        return true;
                    });
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "STA Threading: Failed to execute operation on STA thread");
                    return false;
                }
                
                if (!operationSucceeded)
                {
                    Log.Error("STA Threading: Operation did not complete successfully");
                    return false;
                }
                
                Log.Information("STA Threading: All tests passed successfully");
                return true;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "STA Threading: Test failed with exception");
                return false;
            }
        }
        
        private static async Task<(bool success, string error)> TestQueueOperations()
        {
            try
            {
                var configService = new ConfigurationService();
                var pdfCacheService = new PdfCacheService();
                var processManager = new ProcessManager();
                var officeTracker = new OfficeInstanceTracker();
                
                using var queueService = new SaveQuotesQueueService(configService, pdfCacheService, processManager, officeTracker);
                
                // Create a test PDF file (since PDFs don't require Office)
                var testFilePath = Path.Combine(Path.GetTempPath(), "diagnostic_test.pdf");
                await File.WriteAllTextAsync(testFilePath, "%PDF-1.4\n1 0 obj\n<<\n/Type /Catalog\n/Pages 2 0 R\n>>\nendobj\n2 0 obj\n<<\n/Type /Pages\n/Kids [3 0 R]\n/Count 1\n>>\nendobj\n3 0 obj\n<<\n/Type /Page\n/Parent 2 0 R\n/MediaBox [0 0 612 792]\n>>\nendobj\nxref\n0 4\n0000000000 65535 f \n0000000010 00000 n \n0000000053 00000 n \n0000000125 00000 n \ntrailer\n<<\n/Size 4\n/Root 1 0 R\n>>\nstartxref\n196\n%%EOF");
                
                var fileItem = new FileItem
                {
                    FilePath = testFilePath,
                    FileName = "diagnostic_test.pdf"
                };
                
                // Add to queue
                queueService.AddToQueue(fileItem, "Test Scope", "Test Company", Path.GetTempPath());
                
                if (queueService.TotalCount != 1)
                {
                    return (false, "Failed to add item to queue");
                }
                
                // Clean up test file
                if (File.Exists(testFilePath))
                {
                    File.Delete(testFilePath);
                }
                
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }
    }
} 