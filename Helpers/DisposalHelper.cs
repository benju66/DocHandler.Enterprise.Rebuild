using System;
using System.Runtime.InteropServices;
using Serilog;

namespace DocHandler.Helpers
{
    /// <summary>
    /// Helper methods for safe COM object disposal
    /// </summary>
    public static class DisposalHelper
    {
        private static readonly ILogger _logger = Log.ForContext(typeof(DisposalHelper));
        
        /// <summary>
        /// Safely releases a COM object and sets reference to null
        /// </summary>
        public static void SafeRelease<T>(ref T comObject) where T : class
        {
            if (comObject != null)
            {
                try
                {
                    if (Marshal.IsComObject(comObject))
                    {
                        Marshal.ReleaseComObject(comObject);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Warning(ex, "Error releasing COM object of type {Type}", typeof(T).Name);
                }
                finally
                {
                    comObject = null;
                }
            }
        }
        
        /// <summary>
        /// Forces full garbage collection for COM cleanup
        /// </summary>
        public static void ForceGarbageCollection()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
        
        /// <summary>
        /// Performs aggressive memory cleanup
        /// </summary>
        public static void AggressiveCleanup()
        {
            // Collect all generations
            GC.Collect(2, GCCollectionMode.Forced);
            GC.WaitForPendingFinalizers();
            GC.Collect(2, GCCollectionMode.Forced);
            
            // Compact large object heap
            if (System.Runtime.GCSettings.LargeObjectHeapCompactionMode != System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce)
            {
                System.Runtime.GCSettings.LargeObjectHeapCompactionMode = System.Runtime.GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect();
            }
            
            _logger.Debug("Aggressive memory cleanup completed");
        }
    }
} 