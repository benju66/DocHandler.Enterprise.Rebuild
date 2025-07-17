using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using DocHandler.Services;
using Serilog;

namespace DocHandler
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            // Ensure STA thread for WPF - critical for COM operations
            Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            
            // Run async main method synchronously to preserve STA context
            MainAsync(args).GetAwaiter().GetResult();
        }

        private static async Task MainAsync(string[] args)
        {
            Console.WriteLine("DocHandler Enterprise starting...");
            
            try
            {
                // Initialize logging first
                var logPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DocHandler", "Logs", "log-.txt");
                Console.WriteLine($"Log path: {logPath}");
                
                Log.Logger = new LoggerConfiguration()
                    .WriteTo.Console()
                    .WriteTo.File(
                        path: logPath,
                        rollingInterval: RollingInterval.Day,
                        retainedFileCountLimit: 7,
                        fileSizeLimitBytes: 10 * 1024 * 1024,
                        rollOnFileSizeLimit: true)
                    .CreateLogger();

                Log.Information("DocHandler Enterprise starting...");
                Console.WriteLine("Logger initialized successfully");
                
                // Check for command line arguments
                if (args.Contains("--diagnostic"))
                {
                    Console.WriteLine("Running diagnostic mode...");
                    Log.Information("Running diagnostic mode from command line");
                    
                    try
                    {
                        Console.WriteLine("Starting queue diagnostic...");
                        var diagnosticResult = await QuickDiagnostic.RunQueueDiagnosticAsync();
                        
                        Console.WriteLine("\n" + diagnosticResult);
                        
                        Console.WriteLine("\nDiagnostic completed.");
                        
                        // Try to wait for key, but don't fail if console is redirected
                        try
                        {
                            if (Console.IsInputRedirected == false && Console.KeyAvailable == false)
                            {
                                Console.WriteLine("Press any key to exit...");
                                Console.ReadKey();
                            }
                        }
                        catch (InvalidOperationException)
                        {
                            // Console is redirected, just wait a moment
                            await Task.Delay(1000);
                        }
                        
                        return;
                    }
                    catch (Exception diagEx)
                    {
                        Console.WriteLine($"Diagnostic failed: {diagEx.Message}");
                        Console.WriteLine($"Stack trace: {diagEx.StackTrace}");
                        Log.Error(diagEx, "Diagnostic execution failed");
                        Environment.Exit(1);
                        return;
                    }
                }
                
                // Ensure single instance (optional)
                var mutex = new Mutex(true, "DocHandlerEnterprise", out bool createdNew);
                
                if (!createdNew)
                {
                    Log.Information("Another instance is already running");
                    MessageBox.Show(
                        "DocHandler Enterprise is already running.",
                        "Already Running",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                Console.WriteLine("Single instance check passed");

                try
                {
                    // Verify we're still on STA thread after async operations
                    if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                    {
                        Log.Error("Thread is not STA after async operations - this should not happen");
                        throw new InvalidOperationException("Main thread lost STA state during async operations");
                    }
                    
                    Console.WriteLine($"Thread apartment state verified: {Thread.CurrentThread.GetApartmentState()}");
                    
                    // Create dependency injection container
                    Console.WriteLine("Creating service collection...");
                    var services = new ServiceCollection();
                    
                    Console.WriteLine("Registering services...");
                    services.RegisterServices();
                    
                    Console.WriteLine("Building service provider...");
                    var serviceProvider = services.BuildServiceProvider();
                    
                    Log.Information("Services registered successfully");
                    Console.WriteLine("Services registered successfully");

                    // Create and configure the WPF application
                    Console.WriteLine("Creating WPF application...");
                    var app = new App();
                    
                    Console.WriteLine("Initializing application...");
                    app.InitializeComponent();
                    
                    // Create main window with DI
                    Console.WriteLine("Creating main window...");
                    var mainWindow = new MainWindow(serviceProvider);
                    
                    Console.WriteLine("Setting main window...");
                    app.MainWindow = mainWindow;
                    
                    Console.WriteLine("Showing main window...");
                    mainWindow.Show();
                    
                    Log.Information("Application started successfully");
                    Console.WriteLine("Application started successfully - running message loop");
                    
                    // Run the application
                    app.Run();
                    
                    // Cleanup
                    serviceProvider.Dispose();
                }
                finally
                {
                    mutex?.ReleaseMutex();
                    mutex?.Dispose();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"FATAL ERROR: {ex}");
                Log.Fatal(ex, "Application startup failed");
                
                MessageBox.Show(
                    $"Application failed to start:\n\n{ex.Message}\n\nSee logs for details.",
                    "Startup Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
            finally
            {
                Console.WriteLine("Closing application...");
                Log.CloseAndFlush();
            }
        }
    }
} 