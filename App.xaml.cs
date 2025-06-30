using System;
using System.IO;
using System.Windows;
using Serilog;

namespace DocHandler
{
    public partial class App : Application
    {
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
            
            base.OnStartup(e);
        }
        
        protected override void OnExit(ExitEventArgs e)
        {
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