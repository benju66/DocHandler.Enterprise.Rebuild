using System;
using System.Diagnostics;
using System.IO;
using System.Windows;
using Serilog;

namespace DocHandler
{
    public partial class App : Application
    {
        private System.Windows.Threading.DispatcherTimer _performanceTimer;
        private System.Windows.Threading.DispatcherTimer _memoryTimer;
        protected override void OnStartup(StartupEventArgs e)
        {
            // Initialize Serilog
            var logPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DocHandler",
                "Logs",
                "dochandler-.log");
            
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.File(logPath, 
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fff}] [{Level:u3}] {SourceContext} - {Message:lj}{NewLine}{Exception}")
                .CreateLogger();
            
            Log.Information("DocHandler starting up");
            
            // Set ModernWpfUI theme
            ModernWpf.ThemeManager.Current.ApplicationTheme = ModernWpf.ApplicationTheme.Light;
            
            // Handle unhandled exceptions
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            
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
            // Stop monitoring timers
            _performanceTimer?.Stop();
            _memoryTimer?.Stop();
            
            // Log final performance metrics
            try
            {
                var process = Process.GetCurrentProcess();
                var runtime = DateTime.Now - process.StartTime;
                Log.Information("Application shutdown after {Runtime:hh\\:mm\\:ss} - Peak memory: {PeakMemory}MB",
                               runtime, process.PeakWorkingSet64 / (1024 * 1024));
            }
            catch { }
            
            Log.Information("DocHandler shutting down");
            Log.CloseAndFlush();
            base.OnExit(e);
        }
        
        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Log.Fatal(e.ExceptionObject as Exception, "Unhandled exception occurred");
            
            MessageBox.Show(
                "An unexpected error occurred. The application will now close.\n\nPlease check the log file for details.",
                "Fatal Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }
        
        private void OnDispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Log.Error(e.Exception, "Unhandled dispatcher exception occurred");
            
            MessageBox.Show(
                $"An error occurred: {e.Exception.Message}\n\nThe application will try to continue.",
                "Error",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            
            e.Handled = true;
        }
    }
}