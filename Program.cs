using System;
using System.Threading;
using System.Windows;

namespace DocHandler
{
    public class Program
    {
        [STAThread]
        public static void Main()
        {
            // Ensure single instance (optional)
            var mutex = new Mutex(true, "DocHandlerEnterprise", out bool createdNew);
            
            if (!createdNew)
            {
                MessageBox.Show(
                    "DocHandler Enterprise is already running.",
                    "Already Running",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }
            
            try
            {
                var app = new App();
                app.InitializeComponent();
                app.Run();
            }
            finally
            {
                mutex?.ReleaseMutex();
                mutex?.Dispose();
            }
        }
    }
} 