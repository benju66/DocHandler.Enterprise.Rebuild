using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using DocHandler.ViewModels;
using Serilog;

namespace DocHandler
{
    public partial class App : Application
    {
        private ILogger _logger;
        private System.Windows.Threading.DispatcherTimer _performanceTimer;
        private System.Windows.Threading.DispatcherTimer _memoryTimer;
        protected override void OnStartup(StartupEventArgs e)
        {
            // Initialize logging
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(
                    Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "DocHandler", "Logs", "log-.txt"),
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7)
                .CreateLogger();
                
            _logger = Log.ForContext<App>();
            
            // Set up global exception handlers
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
            
            _logger.Information("DocHandler Enterprise starting up");
            
            // Set ModernWpfUI theme
            ModernWpf.ThemeManager.Current.ApplicationTheme = ModernWpf.ApplicationTheme.Light;
            
            // Performance monitoring timer - log metrics every 5 minutes
            _performanceTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(5)
            };
            _performanceTimer.Tick += (s, args) =>
            {
                try
                {
                    var process = Process.GetCurrentProcess();
                    var workingSet = process.WorkingSet64 / (1024 * 1024); // MB
                    var gcMemory = GC.GetTotalMemory(false) / (1024 * 1024); // MB
                    
                    Log.Information("Performance Check - Memory: {WorkingSet}MB (Working Set), {GcMemory}MB (GC), " +
                                   "Threads: {ThreadCount}, Handles: {HandleCount}",
                                   workingSet, gcMemory, process.Threads.Count, process.HandleCount);
                }
                catch (Exception ex)
                {
                    Log.Warning(ex, "Failed to log performance metrics");
                }
            };
            _performanceTimer.Start();

            // Memory cleanup timer - runs every 30 minutes
            _memoryTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(30)
            };
            _memoryTimer.Tick += (s, args) =>
            {
                Log.Debug("Performing scheduled memory cleanup");
                
                // Force cleanup of generation 2 objects
                GC.Collect(2, GCCollectionMode.Forced);
                GC.WaitForPendingFinalizers();
                GC.Collect(2, GCCollectionMode.Forced);
                
                // Compact large object heap
                System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
                
                Log.Debug("Memory cleanup completed");
            };
            _memoryTimer.Start();
            
            base.OnStartup(e);
        }
        
        protected override void OnExit(ExitEventArgs e)
        {
            _logger.Information("DocHandler Enterprise shutting down");
            
            // Ensure proper cleanup
            if (MainWindow?.DataContext is MainViewModel mainViewModel)
            {
                mainViewModel.Cleanup();
            }
            
            // Stop monitoring timers
            _performanceTimer?.Stop();
            _memoryTimer?.Stop();
            
            // Flush logs
            Log.CloseAndFlush();
            
            base.OnExit(e);
        }
        
        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception;
            _logger.Fatal(exception, "Unhandled exception occurred");
            
            MessageBox.Show(
                $"A fatal error occurred:\n\n{exception?.Message}\n\nThe application will now close.",
                "Fatal Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        
        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            _logger.Error(e.Exception, "Dispatcher unhandled exception");
            
            // Handle specific exceptions gracefully
            if (e.Exception is COMException comEx)
            {
                _logger.Error("COM Exception: HResult={HResult}", comEx.HResult);
                
                MessageBox.Show(
                    "An error occurred with Microsoft Office. The operation will be retried.",
                    "Office Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                    
                e.Handled = true;
            }
            else if (e.Exception is TimeoutException)
            {
                MessageBox.Show(
                    "The operation timed out. Please try again with a smaller file.",
                    "Timeout",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                    
                e.Handled = true;
            }
        }
        
        private void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            _logger.Error(e.Exception, "Unobserved task exception");
            e.SetObserved();
        }
    }
}