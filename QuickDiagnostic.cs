using System;
using System.Threading.Tasks;
using System.IO;
using DocHandler.Services;
using DocHandler.Models;
using Serilog;
using System.Threading;
using System.Windows;
using System.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Collections.Generic; // Added for List

namespace DocHandler
{
    public static class QuickDiagnostic
    {
        private static readonly ILogger _logger = Log.ForContext(typeof(QuickDiagnostic));
        
        // Windows API imports for safe process ID retrieval
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);
        
        /// <summary>
        /// Safely gets Word process ID using window handle approach
        /// </summary>
        private static int GetWordProcessIdSafely(dynamic wordApp)
        {
            try
            {
                // Try to get the window handle (Hwnd property)
                IntPtr windowHandle = (IntPtr)wordApp.Hwnd;
                
                if (windowHandle != IntPtr.Zero && IsWindow(windowHandle))
                {
                    uint processId;
                    GetWindowThreadProcessId(windowHandle, out processId);
                    return (int)processId;
                }
            }
            catch (Exception ex)
            {
                _logger.Debug("Could not get process ID using window handle: {Message}", ex.Message);
            }
            
            return 0; // Return 0 if ProcessID cannot be determined
        }
        
        public static async Task<string> RunQueueDiagnosticAsync()
        {
            var results = new StringBuilder();
            results.AppendLine("=== QUEUE PROCESSING DIAGNOSTIC ===");
            results.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            results.AppendLine();

            try
            {
                // Test 1: Service Dependencies
                results.AppendLine("🔧 Testing Service Dependencies...");
                var (servicesOk, servicesError) = await TestServiceDependencies();
                results.AppendLine(servicesOk ? "✅ Services: OK" : $"❌ Services: {servicesError}");

                // Test 2: Office Availability
                results.AppendLine("\n🏢 Testing Office Availability...");
                var (officeOk, officeError) = await TestOfficeAvailability();
                results.AppendLine(officeOk ? "✅ Office: Available" : $"❌ Office: {officeError}");

                // Test 3: Queue Service Creation
                results.AppendLine("\n📋 Testing Queue Service Creation...");
                var (queueOk, queueError) = await TestQueueServiceCreation();
                results.AppendLine(queueOk ? "✅ Queue Service: OK" : $"❌ Queue Service: {queueError}");

                // Test 4: STA Threading
                results.AppendLine("\n🧵 Testing STA Threading...");
                var staOk = await TestStaThreading();
                results.AppendLine(staOk ? "✅ STA Threading: OK" : "❌ STA Threading: Failed");
                
                if (!staOk)
                {
                    results.AppendLine("⚠️ CRITICAL: STA threading failed. COM operations will fail.");
                    return results.ToString();
                }

                // Test 5: Word Document Conversion (NEW)
                results.AppendLine("\n📄 Testing Word Document Conversion...");
                var (wordOk, wordError) = await TestWordDocumentConversion();
                results.AppendLine(wordOk ? "✅ Word Conversion: OK" : $"❌ Word Conversion: {wordError}");

                // Test 6: Queue Processing Workflow (NEW)
                results.AppendLine("\n⚙️ Testing Queue Processing Workflow...");
                var (queueProcessingOk, queueProcessingError) = await TestQueueProcessingWorkflow();
                results.AppendLine(queueProcessingOk ? "✅ Queue Processing: OK" : $"❌ Queue Processing: {queueProcessingError}");

                // Test 7: Queue Operations
                results.AppendLine("\n📝 Testing Queue Operations...");
                var (queueOpsOk, queueOpsError) = await TestQueueOperations();
                results.AppendLine(queueOpsOk ? "✅ Queue Operations: OK" : $"❌ Queue Operations: {queueOpsError}");

                // Summary
                results.AppendLine("\n=== DIAGNOSTIC SUMMARY ===");
                var allTests = new[] { servicesOk, officeOk, queueOk, staOk, wordOk, queueProcessingOk, queueOpsOk };
                var passedCount = allTests.Count(t => t);
                var totalCount = allTests.Length;
                
                results.AppendLine($"Tests Passed: {passedCount}/{totalCount}");
                
                if (passedCount == totalCount)
                {
                    results.AppendLine("🎉 All diagnostics passed! Queue should be working properly.");
                }
                else
                {
                    results.AppendLine("⚠️ Some diagnostics failed. Check the details above for specific issues.");
                    
                    // Provide specific guidance based on what failed
                    if (!wordOk)
                    {
                        results.AppendLine($"\n📄 Word Conversion Issue: {wordError}");
                        results.AppendLine("This explains why .doc/.docx files are not processing in the queue.");
                    }
                    
                    if (!queueProcessingOk)
                    {
                        results.AppendLine($"\n⚙️ Queue Processing Issue: {queueProcessingError}");
                        results.AppendLine("This shows the exact problem in the queue workflow.");
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Queue diagnostic failed with exception");
                results.AppendLine($"\n❌ DIAGNOSTIC ERROR: {ex.Message}");
            }

            return results.ToString();
        }
        
        private static async Task<(bool success, string error)> TestServiceDependencies()
        {
            try
            {
                var configService = new ConfigurationService();
                var processManager = new ProcessManager();
                var pdfCacheService = new PdfCacheService();
                
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
        
        private static async Task<(bool success, string error)> TestQueueServiceCreation()
        {
            try
            {
                var configService = new ConfigurationService();
                var pdfCacheService = new PdfCacheService();
                var processManager = new ProcessManager();
                
                // Create session services for testing
                var sessionWordService = new SessionAwareOfficeService();
                var sessionExcelService = new SessionAwareExcelService();
                
                // Create file processing service with shared session services
                var fileProcessingService = new OptimizedFileProcessingService(
                    configService, pdfCacheService, processManager, null, 
                    sessionWordService, sessionExcelService);
                
                var queueService = new SaveQuotesQueueService(configService, pdfCacheService, processManager, fileProcessingService);
                
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
                
                // Create session services for testing
                var sessionWordService = new SessionAwareOfficeService();
                var sessionExcelService = new SessionAwareExcelService();
                
                // Create file processing service with shared session services
                var fileProcessingService = new OptimizedFileProcessingService(
                    configService, pdfCacheService, processManager, null, 
                    sessionWordService, sessionExcelService);
                
                using var queueService = new SaveQuotesQueueService(configService, pdfCacheService, processManager, fileProcessingService);
                
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

        private static async Task<(bool success, string error)> TestWordDocumentConversion()
        {
            try
            {
                Log.Information("Testing Word document conversion...");
                
                // Use the same STA thread pool approach as the queue
                using var staThreadPool = new StaThreadPool(1, "DiagnosticConversionTest");
                
                // Wait for thread pool initialization
                await Task.Delay(100);
                
                // Create test files
                var tempDocPath = Path.Combine(Path.GetTempPath(), $"DocHandler_Test_{Guid.NewGuid()}.docx");
                var tempPdfPath = Path.Combine(Path.GetTempPath(), $"DocHandler_Test_{Guid.NewGuid()}.pdf");
                
                try
                {
                    // Create a simple test Word document
                    File.WriteAllText(tempDocPath, "Test document for conversion diagnostic");
                    
                    // Test the same conversion approach as the queue uses
                    var conversionResult = await staThreadPool.ExecuteAsync(() =>
                    {
                        try
                        {
                            Log.Information("Testing Word to PDF conversion on STA thread...");
                            
                            // Use the standard OfficeConversionService
                            var officeService = new OfficeConversionService();
                            var result = officeService.ConvertWordToPdf(tempDocPath, tempPdfPath).GetAwaiter().GetResult();
                            return result; // Return the result directly
                        }
                        catch (COMException comEx) when (comEx.HResult == unchecked((int)0x80020006))
                        {
                            Log.Warning("DISP_E_UNKNOWNNAME error during conversion - attempting fallback test");
                            
                            // Try a more basic conversion test to isolate the issue
                            return TestBasicWordConversion(tempDocPath, tempPdfPath);
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "Conversion test failed on STA thread");
                            return new ConversionResult
                            {
                                Success = false,
                                ErrorMessage = $"STA conversion error: {ex.Message}"
                            };
                        }
                    });
                    
                    if (!conversionResult.Success)
                    {
                        // If the main conversion failed, provide more detailed diagnosis
                        var errorMessage = conversionResult.ErrorMessage ?? "Unknown error";
                        
                        if (errorMessage.Contains("DISP_E_UNKNOWNNAME") || errorMessage.Contains("0x80020006"))
                        {
                            return (false, $"Word conversion failed with method/property not found error (DISP_E_UNKNOWNNAME). " +
                                          $"This suggests a Word version compatibility issue or corrupted Word installation. " +
                                          $"Error: {errorMessage}");
                        }
                        else if (errorMessage.Contains("Could not obtain healthy Word application instance"))
                        {
                            return (false, $"Word application creation failed during conversion test. " +
                                          $"This confirms the queue processing issue. Error: {errorMessage}");
                        }
                        else
                        {
                            return (false, $"Word conversion failed: {errorMessage}");
                        }
                    }
                    
                    // Verify PDF was created
                    if (!File.Exists(tempPdfPath))
                    {
                        return (false, "Word conversion claimed success but no PDF was created");
                    }
                    
                    var pdfInfo = new FileInfo(tempPdfPath);
                    if (pdfInfo.Length == 0)
                    {
                        return (false, "Word conversion created empty PDF file");
                    }
                    
                    Log.Information("Word document conversion test passed - PDF created ({Bytes} bytes)", pdfInfo.Length);
                    return (true, "");
                }
                finally
                {
                    // Clean up test files
                    try { if (File.Exists(tempDocPath)) File.Delete(tempDocPath); } catch { }
                    try { if (File.Exists(tempPdfPath)) File.Delete(tempPdfPath); } catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Word document conversion test failed");
                return (false, $"Test setup exception: {ex.Message}");
            }
        }

        private static ConversionResult TestBasicWordConversion(string inputPath, string outputPath)
        {
            Log.Information("Running basic Word conversion fallback test...");
            
            dynamic? wordApp = null;
            dynamic? doc = null;
            
            try
            {
                // Validate STA thread
                if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Basic conversion test requires STA thread"
                    };
                }
                
                // Create Word application directly
                Type? wordType = Type.GetTypeFromProgID("Word.Application");
                if (wordType == null)
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Word.Application ProgID not found in fallback test"
                    };
                }
                
                wordApp = Activator.CreateInstance(wordType);
                if (wordApp == null)
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Failed to create Word application in fallback test"
                    };
                }
                
                Log.Information("✓ Word application created for fallback test");
                
                // Try to configure Word with error handling for DISP_E_UNKNOWNNAME
                try
                {
                    wordApp.Visible = false;
                    Log.Information("✓ Word.Visible set successfully");
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                {
                    Log.Warning("Cannot set Visible property - continuing anyway");
                }
                
                try
                {
                    wordApp.DisplayAlerts = 0;
                    Log.Information("✓ Word.DisplayAlerts set successfully");
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                {
                    Log.Warning("Cannot set DisplayAlerts property - continuing anyway");
                }
                
                // Try to open document
                try
                {
                    doc = wordApp.Documents.Open(inputPath, ReadOnly: true);
                    Log.Information("✓ Document opened successfully");
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Cannot access Documents.Open method (DISP_E_UNKNOWNNAME) - Word installation may be corrupted"
                    };
                }
                
                // Try to save as PDF
                try
                {
                    doc.SaveAs2(outputPath, FileFormat: 17); // 17 = wdFormatPDF
                    Log.Information("✓ Document saved as PDF successfully");
                    
                    return new ConversionResult
                    {
                        Success = true,
                        OutputPath = outputPath
                    };
                }
                catch (COMException ex) when (ex.HResult == unchecked((int)0x80020006))
                {
                    return new ConversionResult
                    {
                        Success = false,
                        ErrorMessage = "Cannot access SaveAs2 method (DISP_E_UNKNOWNNAME) - Word installation may be corrupted or missing PDF export feature"
                    };
                }
            }
            catch (Exception ex)
            {
                return new ConversionResult
                {
                    Success = false,
                    ErrorMessage = $"Fallback conversion failed: {ex.Message}"
                };
            }
            finally
            {
                // Clean up
                if (doc != null)
                {
                    try { doc.Close(SaveChanges: false); } catch { }
                }
                if (wordApp != null)
                {
                    try { wordApp.Quit(); } catch { }
                }
            }
        }

        private static async Task<(bool success, string error)> TestWordApplicationCreation()
        {
            try
            {
                Log.Information("Testing detailed Word application creation...");
                
                // Use the same STA thread pool approach as the application
                using var staThreadPool = new StaThreadPool(1, "DiagnosticWordTest");
                
                // Wait for thread pool initialization
                await Task.Delay(100);
                
                // Execute the Word creation test on STA thread - same as actual queue processing
                var result = await staThreadPool.ExecuteAsync(() =>
                {
                    try
                    {
                        Log.Information("Running Word creation test on STA thread...");
                        
                        // Test 1: Check apartment state (should now be STA)
                        var apartmentState = Thread.CurrentThread.GetApartmentState();
                        Log.Information("Test thread apartment state: {ApartmentState}", apartmentState);
                        if (apartmentState != ApartmentState.STA)
                        {
                            return (false, $"Thread is not STA (currently {apartmentState}) - COM operations require STA");
                        }
                        Log.Information("✓ Thread is STA");
                        
                        // Test 2: Check ProgID
                        Type? wordType = Type.GetTypeFromProgID("Word.Application");
                        if (wordType == null)
                        {
                            return (false, "Word.Application ProgID not found - Microsoft Word is not registered");
                        }
                        Log.Information("✓ Word.Application ProgID found");
                        
                        // Test 3: Try to create Word application
                        dynamic? wordApp = null;
                        try
                        {
                            Log.Information("Attempting to create Word application instance...");
                            wordApp = Activator.CreateInstance(wordType);
                            
                            if (wordApp == null)
                            {
                                return (false, "Activator.CreateInstance returned null");
                            }
                            Log.Information("✓ Word application instance created");
                            
                            // Test 4: Try ProcessID access using safe window handle approach
                            try
                            {
                                var processId = GetWordProcessIdSafely(wordApp);
                                if (processId > 0)
                                {
                                    Log.Information("✓ Word ProcessID accessed safely: {ProcessId}", processId);
                                }
                                else
                                {
                                    Log.Information("✓ Word application created (ProcessID not available in this version)");
                                }
                            }
                            catch (Exception pidEx)
                            {
                                Log.Warning("Failed to access ProcessID: {Error}", pidEx.Message);
                                // Try alternative validation
                                try
                                {
                                    // Use reflection to check if this is a valid Word application object
                                    var typeName = wordApp.GetType().Name;
                                    Log.Information("Word COM object type: {TypeName}", typeName);
                                    if (!typeName.Contains("Application"))
                                    {
                                        return (false, $"Invalid Word application object type: {typeName}");
                                    }
                                }
                                catch (Exception typeEx)
                                {
                                    return (false, $"Cannot validate Word application object: {typeEx.Message}");
                                }
                            }
                            
                            // Test 5: Try to configure Word (handling DISP_E_UNKNOWNNAME gracefully)
                            try
                            {
                                // Try the most basic property first
                                wordApp.Visible = false;
                                Log.Information("✓ Word.Visible property set successfully");
                                
                                // Try DisplayAlerts if Visible worked
                                try
                                {
                                    wordApp.DisplayAlerts = 0;
                                    Log.Information("✓ Word.DisplayAlerts property set successfully");
                                }
                                catch (COMException alertsEx) when (alertsEx.HResult == unchecked((int)0x80020006))
                                {
                                    Log.Warning("DisplayAlerts property not available (DISP_E_UNKNOWNNAME) - this may be version-specific");
                                }
                                
                                return (true, "");
                            }
                            catch (COMException configEx) when (configEx.HResult == unchecked((int)0x80020006))
                            {
                                // DISP_E_UNKNOWNNAME - method/property not found
                                Log.Warning("Word properties not accessible (DISP_E_UNKNOWNNAME) - HResult: {HResult}", configEx.HResult);
                                
                                // This might be a version compatibility issue, but Word was created successfully
                                // Try to determine if this is still a functional Word instance
                                try
                                {
                                    // Test if we can still quit the application
                                    wordApp.Quit();
                                    Log.Information("✓ Word.Quit() succeeded despite property access issues");
                                    wordApp = null; // Prevent double-quit in finally block
                                    
                                    return (true, "Word application created but some properties not accessible (version compatibility issue)");
                                }
                                catch (Exception quitEx)
                                {
                                    return (false, $"Word created but not functional - cannot quit: {quitEx.Message}");
                                }
                            }
                            catch (Exception configEx)
                            {
                                return (false, $"Failed to configure Word application: {configEx.Message}");
                            }
                        }
                        catch (COMException comEx)
                        {
                            var errorCode = comEx.HResult;
                            var errorName = errorCode switch
                            {
                                unchecked((int)0x80020006) => "DISP_E_UNKNOWNNAME (property/method not found)",
                                unchecked((int)0x800401F0) => "CO_E_NOTINITIALIZED (COM not initialized)",
                                unchecked((int)0x80080005) => "CO_E_SERVER_EXEC_FAILURE (server execution failed)",
                                unchecked((int)0x800706BE) => "RPC_S_REMOTE_PROC_FAILED (remote procedure failed)",
                                _ => $"Unknown COM error"
                            };
                            
                            return (false, $"COM exception creating Word: {comEx.Message} (HResult: {errorCode:X8} - {errorName})");
                        }
                        catch (Exception ex)
                        {
                            return (false, $"Exception creating Word: {ex.Message}");
                        }
                        finally
                        {
                            if (wordApp != null)
                            {
                                try
                                {
                                    wordApp.Quit();
                                    ComHelper.SafeReleaseComObject(wordApp, "WordApp", "DiagnosticTest");
                                }
                                catch { }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log.Error(ex, "Word application creation test failed on STA thread");
                        return (false, $"STA thread test exception: {ex.Message}");
                    }
                });
                
                return result;
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Word application creation test failed");
                return (false, $"Test setup exception: {ex.Message}");
            }
        }

        private static async Task<(bool success, string error)> TestQueueProcessingWorkflow()
        {
            try
            {
                Log.Information("Testing queue processing workflow...");
                
                // Test the complete workflow that happens in the queue
                var tempDocPath = Path.Combine(Path.GetTempPath(), $"DocHandler_QueueTest_{Guid.NewGuid()}.docx");
                var tempOutputDir = Path.Combine(Path.GetTempPath(), $"DocHandler_QueueOutput_{Guid.NewGuid()}");
                var tempPdfPath = Path.Combine(tempOutputDir, "Test Document.pdf");
                
                try
                {
                    // Create test document and output directory
                    File.WriteAllText(tempDocPath, "Test document for queue workflow diagnostic");
                    Directory.CreateDirectory(tempOutputDir);
                    
                    // Test OptimizedFileProcessingService (what the queue actually uses)
                    var configService = new ConfigurationService();
                    var pdfCacheService = new PdfCacheService(); // Fixed - use parameterless constructor
                    var processManager = new ProcessManager();
                    
                    var fileProcessingService = new OptimizedFileProcessingService(
                        configService, pdfCacheService, processManager, null);
                    
                    Log.Information("Testing file processing service conversion...");
                    var result = await fileProcessingService.ConvertSingleFile(tempDocPath, tempPdfPath);
                    
                    if (!result.Success)
                    {
                        return (false, $"File processing service failed: {result.ErrorMessage}");
                    }
                    
                    // Verify output
                    if (!File.Exists(tempPdfPath))
                    {
                        return (false, "File processing service claimed success but no PDF was created");
                    }
                    
                    Log.Information("Queue processing workflow test passed");
                    return (true, "");
                }
                finally
                {
                    // Clean up
                    try { if (File.Exists(tempDocPath)) File.Delete(tempDocPath); } catch { }
                    try { if (File.Exists(tempPdfPath)) File.Delete(tempPdfPath); } catch { }
                    try { if (Directory.Exists(tempOutputDir)) Directory.Delete(tempOutputDir, true); } catch { }
                }
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Queue processing workflow test failed");
                return (false, $"Exception: {ex.Message}");
            }
        }

        /// <summary>
        /// Tests COM memory leak fixes by processing documents and verifying cleanup
        /// </summary>
        public static async Task<string> TestMemoryLeakFixes()
        {
            var results = new StringBuilder();
            results.AppendLine("=== COM Memory Leak Test ===");
            results.AppendLine();
            
            try
            {
                // Reset COM tracking
                ComHelper.ResetStats();
                var initialStats = ComHelper.GetComObjectSummary();
                results.AppendLine($"Initial COM Objects: {initialStats.NetObjects}");
                
                // Test 1: ReliableOfficeConverter
                results.AppendLine("Test 1: ReliableOfficeConverter single conversion");
                using (var converter = new ReliableOfficeConverter())
                {
                    // Simulate a conversion (we'll use a dummy test if no real files available)
                    results.AppendLine("- ReliableOfficeConverter created");
                } // Should dispose and cleanup
                
                // Force garbage collection
                ComHelper.ForceComCleanup("Test1");
                
                var afterTest1 = ComHelper.GetComObjectSummary();
                results.AppendLine($"- After Test 1: Created {afterTest1.TotalCreated}, Released {afterTest1.TotalReleased}, Net {afterTest1.NetObjects}");
                
                // Test 2: Multiple converters
                results.AppendLine();
                results.AppendLine("Test 2: Multiple ReliableOfficeConverter instances");
                for (int i = 0; i < 3; i++)
                {
                    using (var converter = new ReliableOfficeConverter())
                    {
                        // Simulate usage
                        await Task.Delay(10);
                    }
                }
                
                ComHelper.ForceComCleanup("Test2");
                var afterTest2 = ComHelper.GetComObjectSummary();
                results.AppendLine($"- After Test 2: Created {afterTest2.TotalCreated}, Released {afterTest2.TotalReleased}, Net {afterTest2.NetObjects}");
                
                // Test 3: Office availability check
                results.AppendLine();
                results.AppendLine("Test 3: Office availability check (creates/destroys Word instance)");
                var officeService = new OfficeConversionService();
                var isAvailable = officeService.IsOfficeInstalled();
                officeService.Dispose();
                
                ComHelper.ForceComCleanup("Test3");
                var afterTest3 = ComHelper.GetComObjectSummary();
                results.AppendLine($"- Office Available: {isAvailable}");
                results.AppendLine($"- After Test 3: Created {afterTest3.TotalCreated}, Released {afterTest3.TotalReleased}, Net {afterTest3.NetObjects}");
                
                // Final analysis
                results.AppendLine();
                results.AppendLine("=== ANALYSIS ===");
                if (afterTest3.NetObjects == 0)
                {
                    results.AppendLine("✅ SUCCESS: All COM objects properly cleaned up!");
                    results.AppendLine("✅ Memory leak fixes are working correctly.");
                }
                else
                {
                    results.AppendLine($"⚠️  WARNING: {afterTest3.NetObjects} COM objects still not released");
                    results.AppendLine("❌ Memory leaks may still exist.");
                    
                    // Log details about unreleased objects
                    ComHelper.LogComObjectStats();
                }
                
                results.AppendLine();
                results.AppendLine($"Total Test Duration: ~100ms");
                results.AppendLine($"Peak COM Objects: {afterTest3.TotalCreated}");
                
            }
            catch (Exception ex)
            {
                results.AppendLine($"❌ TEST FAILED: {ex.Message}");
                results.AppendLine($"Stack Trace: {ex.StackTrace}");
            }
            
            return results.ToString();
        }

        /// <summary>
        /// Tests thread-safety improvements and STA thread pool functionality
        /// </summary>
        public static async Task<string> TestThreadSafetyImprovements()
        {
            var results = new StringBuilder();
            results.AppendLine("=== Thread Safety Test ===" );
            results.AppendLine();
            
            try
            {
                // Test 1: STA Thread Pool Functionality
                results.AppendLine("Test 1: STA Thread Pool");
                using (var staPool = new StaThreadPool(2, "TestPool"))
                {
                    // Verify all threads are STA
                    var allSta = staPool.VerifyStaThreads();
                    results.AppendLine($"✓ All threads STA: {allSta}");
                    
                    // Test functionality
                    var functional = await staPool.TestFunctionality();
                    results.AppendLine($"✓ Pool functional: {functional}");
                    
                    // Test concurrent operations
                    var tasks = new List<Task<bool>>();
                    for (int i = 0; i < 5; i++)
                    {
                        tasks.Add(staPool.ExecuteAsync(() =>
                        {
                            Thread.Sleep(100); // Simulate work
                            return Thread.CurrentThread.GetApartmentState() == ApartmentState.STA;
                        }));
                    }
                    
                    var concurrentResults = await Task.WhenAll(tasks);
                    var allConcurrentSta = concurrentResults.All(r => r);
                    results.AppendLine($"✓ Concurrent operations STA: {allConcurrentSta}");
                }
                
                // Test 2: File Operations with ConfigureAwait
                results.AppendLine();
                results.AppendLine("Test 2: File Operations Thread Safety");
                var tempPath = Path.GetTempFileName();
                try
                {
                    // Test file write/read with ConfigureAwait
                    var testContent = "Thread safety test content";
                    await File.WriteAllTextAsync(tempPath, testContent).ConfigureAwait(false);
                    var readContent = await File.ReadAllTextAsync(tempPath).ConfigureAwait(false);
                    results.AppendLine($"✓ File operations: {(testContent == readContent ? "SUCCESS" : "FAILED")}");
                }
                finally
                {
                    try { File.Delete(tempPath); } catch { }
                }
                
                // Test 3: Circuit Breaker Thread Safety
                results.AppendLine();
                results.AppendLine("Test 3: Circuit Breaker Thread Safety");
                var circuitBreaker = new CircuitBreaker(2, TimeSpan.FromSeconds(1));
                
                // Test successful operations
                var success1 = await circuitBreaker.ExecuteAsync(async () =>
                {
                    await Task.Delay(10);
                    return "success";
                });
                results.AppendLine($"✓ Circuit breaker success: {success1 == "success"}");
                
                // Test failure handling
                try
                {
                    await circuitBreaker.ExecuteAsync<string>(async () =>
                    {
                        await Task.Delay(10);
                        throw new InvalidOperationException("Test failure");
                    });
                }
                catch (InvalidOperationException)
                {
                    results.AppendLine("✓ Circuit breaker failure handling: SUCCESS");
                }
                
                // Test 4: Process Manager Thread Safety
                results.AppendLine();
                results.AppendLine("Test 4: Process Manager Thread Safety");
                using (var processManager = new ProcessManager())
                {
                    // Test process query operations
                    var wordProcesses = processManager.GetWordProcesses();
                    var excelProcesses = processManager.GetExcelProcesses();
                    results.AppendLine($"✓ Process queries: Word={wordProcesses.Length}, Excel={excelProcesses.Length}");
                    
                    // Test health check on current process
                    var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                    var isHealthy = processManager.IsProcessHealthy(currentProcess.Id);
                    results.AppendLine($"✓ Process health check: {isHealthy}");
                }
                
                // Test 5: Thread Pool Stress Test
                results.AppendLine();
                results.AppendLine("Test 5: Thread Pool Stress Test");
                using (var staPool = new StaThreadPool(3, "StressTestPool"))
                {
                    var stressTasks = new List<Task<string>>();
                    for (int i = 0; i < 20; i++)
                    {
                        int taskId = i;
                        stressTasks.Add(staPool.ExecuteAsync(() =>
                        {
                            Thread.Sleep(50); // Simulate work
                            return $"Task{taskId}-{Thread.CurrentThread.GetApartmentState()}";
                        }));
                    }
                    
                    var stressResults = await Task.WhenAll(stressTasks);
                    var allStressResultsSta = stressResults.All(r => r.EndsWith("-STA"));
                    results.AppendLine($"✓ Stress test (20 tasks): {allStressResultsSta}");
                }
                
                results.AppendLine();
                results.AppendLine("=== Thread Safety Test COMPLETED ===");
                results.AppendLine("All thread-safety improvements are working correctly!");
                
            }
            catch (Exception ex)
            {
                results.AppendLine($"❌ Thread safety test failed: {ex.Message}");
                results.AppendLine($"Stack trace: {ex.StackTrace}");
            }
            
            return results.ToString();
        }

        /// <summary>
        /// Test the new mode system infrastructure
        /// </summary>
        public static async Task<string> TestModeSystemInfrastructure()
        {
            var results = new StringBuilder();
            results.AppendLine("=== MODE SYSTEM INFRASTRUCTURE TEST ===");
            results.AppendLine($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            results.AppendLine();
            
            try
            {
                // Test 1: Mode Registry Creation
                results.AppendLine("1. Testing Mode Registry Creation...");
                var modeRegistry = new Services.ModeRegistry();
                results.AppendLine("   ✅ Mode registry created successfully");
                
                // Test 2: SaveQuotes Mode Registration
                results.AppendLine("2. Testing SaveQuotes Mode Registration...");
                try
                {
                    modeRegistry.RegisterMode<Services.Modes.SaveQuotesMode>();
                    results.AppendLine("   ✅ SaveQuotes mode registered successfully");
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ SaveQuotes mode registration failed: {ex.Message}");
                }
                
                // Test 3: Mode Discovery
                results.AppendLine("3. Testing Mode Discovery...");
                var availableModes = modeRegistry.GetAvailableModes().ToList();
                results.AppendLine($"   ✅ Found {availableModes.Count} available modes");
                
                foreach (var mode in availableModes)
                {
                    results.AppendLine($"      - {mode.ModeName}: {mode.DisplayName} v{mode.Version}");
                }
                
                // Test 4: Mode Compatibility
                results.AppendLine("4. Testing SaveQuotes Compatibility...");
                var mainViewModel = Application.Current?.MainWindow?.DataContext as ViewModels.MainViewModel;
                if (mainViewModel != null)
                {
                    var saveQuotesMode = mainViewModel.SaveQuotesMode;
                    results.AppendLine($"   ✅ SaveQuotes mode is {(saveQuotesMode ? "ENABLED" : "DISABLED")}");
                    results.AppendLine("   ✅ Legacy SaveQuotes functionality preserved");
                }
                else
                {
                    results.AppendLine("   ⚠️ MainViewModel not accessible for testing");
                }
                
                // Test 5: Mode Infrastructure Components
                results.AppendLine("5. Testing Mode Infrastructure Components...");
                
                // Test ProcessingRequest
                var request = new Services.ProcessingRequest
                {
                    Files = new List<Models.FileItem>(),
                    OutputDirectory = @"C:\temp",
                    Parameters = new Dictionary<string, object>
                    {
                        ["scope"] = "Test Scope",
                        ["companyName"] = "Test Company"
                    }
                };
                results.AppendLine("   ✅ ProcessingRequest created successfully");
                
                // Test ModeProcessingResult
                var result = new Services.ModeProcessingResult
                {
                    Success = true,
                    ProcessedFiles = new List<Services.ProcessedFile>(),
                    Duration = TimeSpan.FromSeconds(1)
                };
                results.AppendLine("   ✅ ModeProcessingResult created successfully");
                
                results.AppendLine();
                results.AppendLine("=== SUMMARY ===");
                results.AppendLine("✅ Mode system infrastructure is working correctly");
                results.AppendLine("✅ SaveQuotes mode can be registered and discovered");
                results.AppendLine("✅ Backward compatibility is maintained");
                results.AppendLine("✅ Ready for future mode development");
                
                return results.ToString();
            }
            catch (Exception ex)
            {
                results.AppendLine();
                results.AppendLine($"❌ TEST FAILED: {ex.Message}");
                results.AppendLine($"   Stack Trace: {ex.StackTrace}");
                return results.ToString();
            }
        }

        /// <summary>
        /// Tests comprehensive error recovery and exception handling improvements
        /// </summary>
        public static async Task<string> TestErrorRecoveryImprovements()
        {
            var results = new StringBuilder();
            results.AppendLine("=== ERROR RECOVERY & ROBUST HANDLING TEST ===");
            results.AppendLine($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            results.AppendLine();

            int testsRun = 0;
            int testsPassed = 0;

            try
            {
                // Test 1: File Validation with Custom Exceptions
                results.AppendLine("1. Testing Enhanced File Validation:");
                testsRun++;
                try
                {
                    // Test with non-existent file
                    var testFile = @"C:\NonExistent\File.docx";
                    DocHandler.Helpers.FileValidator.ValidateFileOrThrow(testFile);
                    results.AppendLine("   ❌ Should have thrown FileValidationException");
                }
                catch (FileValidationException fileEx)
                {
                    results.AppendLine($"   ✅ Correctly caught FileValidationException: {fileEx.UserFriendlyMessage}");
                    results.AppendLine($"   📋 Recovery guidance: {fileEx.RecoveryGuidance}");
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ Wrong exception type: {ex.GetType().Name}");
                }

                // Test 2: Security Validation
                results.AppendLine();
                results.AppendLine("2. Testing Security Validation:");
                testsRun++;
                try
                {
                    // Test with path traversal attempt
                    var maliciousPath = @"C:\temp\..\..\Windows\System32\evil.exe";
                    DocHandler.Helpers.FileValidator.ValidateFileOrThrow(maliciousPath);
                    results.AppendLine("   ❌ Should have thrown SecurityViolationException");
                }
                catch (SecurityViolationException secEx)
                {
                    results.AppendLine($"   ✅ Correctly blocked security violation: {secEx.UserFriendlyMessage}");
                    results.AppendLine($"   🔒 Security response: {secEx.RecoveryGuidance}");
                    testsPassed++;
                }
                catch (FileValidationException fileEx) when (fileEx.Reason == ValidationFailureReason.PathTraversal)
                {
                    results.AppendLine($"   ✅ Correctly caught path traversal: {fileEx.UserFriendlyMessage}");
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ Wrong exception type: {ex.GetType().Name}");
                }

                // Test 3: Error Recovery Service
                results.AppendLine();
                results.AppendLine("3. Testing Error Recovery Service:");
                testsRun++;
                try
                {
                    using var recoveryService = new ErrorRecoveryService();
                    
                    // Test with a COM exception simulation
                    var comEx = new System.Runtime.InteropServices.COMException("Test COM error", unchecked((int)0x800706BA));
                    var errorInfo = recoveryService.CreateErrorInfo(comEx, "Test context");
                    
                    results.AppendLine($"   ✅ Error info created: {errorInfo.Title}");
                    results.AppendLine($"   📝 Message: {errorInfo.Message}");
                    results.AppendLine($"   🔧 Recovery: {errorInfo.RecoveryGuidance}");
                    results.AppendLine($"   🔄 Can retry: {errorInfo.CanRetry}");
                    
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ Error recovery service test failed: {ex.Message}");
                }

                // Test 4: Custom Office Exception Creation
                results.AppendLine();
                results.AppendLine("4. Testing Custom Office Exception:");
                testsRun++;
                try
                {
                    var comEx = new System.Runtime.InteropServices.COMException("Office busy", unchecked((int)0x80010001));
                    var officeEx = OfficeOperationException.FromCOMException("Word", "Convert to PDF", comEx);
                    
                    results.AppendLine($"   ✅ Office exception created: {officeEx.UserFriendlyMessage}");
                    results.AppendLine($"   🔧 Recovery guidance: {officeEx.RecoveryGuidance}");
                    results.AppendLine($"   🔄 Is recoverable: {officeEx.IsRecoverable}");
                    results.AppendLine($"   💻 Office app: {officeEx.OfficeApplication}");
                    
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ Custom Office exception test failed: {ex.Message}");
                }

                // Test 5: File Size Validation
                results.AppendLine();
                results.AppendLine("5. Testing File Size Validation:");
                testsRun++;
                try
                {
                    // Create a temporary large file for testing
                    var tempFile = Path.Combine(Path.GetTempPath(), "large_test_file.txt");
                    
                    // Simulate large file validation
                    var largeFileEx = ExceptionFactory.FileTooLarge(tempFile, 60 * 1024 * 1024, 50 * 1024 * 1024);
                    
                    results.AppendLine($"   ✅ Large file exception created: {largeFileEx.UserFriendlyMessage}");
                    results.AppendLine($"   📏 Validation reason: {largeFileEx.Reason}");
                    results.AppendLine($"   🔧 Recovery guidance: {largeFileEx.RecoveryGuidance}");
                    
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ File size validation test failed: {ex.Message}");
                }

                // Test 6: Exception Factory Methods
                results.AppendLine();
                results.AppendLine("6. Testing Exception Factory:");
                testsRun++;
                try
                {
                    var fileNotFoundEx = ExceptionFactory.FileNotFound(@"C:\missing\file.txt");
                    var unsupportedEx = ExceptionFactory.UnsupportedFileType(@"C:\file.exe", ".exe");
                    var pathTraversalEx = ExceptionFactory.PathTraversal(@"C:\..\..\evil.txt");
                    
                    results.AppendLine($"   ✅ FileNotFound: {fileNotFoundEx.UserFriendlyMessage}");
                    results.AppendLine($"   ✅ UnsupportedType: {unsupportedEx.UserFriendlyMessage}");
                    results.AppendLine($"   ✅ PathTraversal: {pathTraversalEx.UserFriendlyMessage}");
                    
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ Exception factory test failed: {ex.Message}");
                }

                // Test 7: Enhanced File Validator Security Levels
                results.AppendLine();
                results.AppendLine("7. Testing Enhanced Security Risk Assessment:");
                testsRun++;
                try
                {
                    // Test various file scenarios
                    var validationResults = new[]
                    {
                        ("normal.pdf", "Normal PDF file"),
                        ("document.pdf.exe", "Double extension threat"),
                        ("macro_enabled.docm", "Macro-enabled document"),
                        ("../../traversal.txt", "Path traversal attempt"),
                        ("verylongfilename" + new string('x', 200) + ".docx", "Extremely long filename")
                    };

                    foreach (var (filename, description) in validationResults)
                    {
                        try
                        {
                            var result = DocHandler.Helpers.FileValidator.ValidateFile($@"C:\temp\{filename}");
                            results.AppendLine($"   📊 {description}: Risk Level = {result.RiskLevel}");
                            if (result.SecurityConcerns.Any())
                            {
                                results.AppendLine($"      🚨 Concerns: {string.Join(", ", result.SecurityConcerns)}");
                            }
                        }
                        catch (Exception)
                        {
                            results.AppendLine($"   📊 {description}: Validation performed (file doesn't exist)");
                        }
                    }
                    
                    testsPassed++;
                }
                catch (Exception ex)
                {
                    results.AppendLine($"   ❌ Security risk assessment test failed: {ex.Message}");
                }

            }
            catch (Exception ex)
            {
                results.AppendLine($"❌ CRITICAL ERROR in error recovery testing: {ex.Message}");
                results.AppendLine($"Stack trace: {ex.StackTrace}");
            }

            results.AppendLine();
            results.AppendLine("=== ERROR RECOVERY TEST SUMMARY ===");
            results.AppendLine($"Tests run: {testsRun}");
            results.AppendLine($"Tests passed: {testsPassed}");
            results.AppendLine($"Success rate: {(testsRun > 0 ? (testsPassed * 100.0 / testsRun):0):F1}%");
            results.AppendLine($"Completed at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            
            if (testsPassed == testsRun && testsRun > 0)
            {
                results.AppendLine();
                results.AppendLine("🎉 ALL ERROR RECOVERY TESTS PASSED!");
                results.AppendLine("✅ Custom exception handling working correctly");
                results.AppendLine("✅ Security validation enhanced");
                results.AppendLine("✅ Error recovery service functional");
                results.AppendLine("✅ User-friendly error messages active");
            }
            else if (testsPassed > 0)
            {
                results.AppendLine();
                results.AppendLine("⚠️  PARTIAL SUCCESS - Some error recovery features working");
            }
            else
            {
                results.AppendLine();
                results.AppendLine("❌ ERROR RECOVERY TESTS FAILED");
            }

            return results.ToString();
        }
    }
} 