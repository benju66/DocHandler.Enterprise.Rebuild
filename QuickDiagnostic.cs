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
    }
} 