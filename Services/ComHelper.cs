using System;
using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Threading;
using Serilog;
using System.Collections.Generic;

namespace DocHandler.Services
{
    /// <summary>
    /// Helper class for tracking COM object lifecycle to detect memory leaks
    /// </summary>
    public static class ComHelper
    {
        private static readonly ILogger _logger = Log.ForContext(typeof(ComHelper));
        private static readonly ConcurrentDictionary<string, ComObjectStats> _stats = new();
        private static long _totalCreated = 0;
        private static long _totalReleased = 0;

        /// <summary>
        /// Safely releases a COM object and tracks the disposal
        /// </summary>
        /// <param name="comObject">The COM object to release</param>
        /// <param name="objectType">Type of object for tracking (e.g., "WordApp", "Document")</param>
        /// <param name="context">Context information for logging</param>
        /// <returns>True if successfully released, false otherwise</returns>
        public static bool SafeReleaseComObject(object? comObject, string objectType, string context = "")
        {
            if (comObject == null)
                return true;

            try
            {
                // Use regular ReleaseComObject, NOT FinalReleaseComObject
                var refCount = Marshal.ReleaseComObject(comObject);
                
                // Track the disposal
                TrackComObjectDisposal(objectType, context);
                
                _logger.Debug("Released COM object {ObjectType} in {Context}, RefCount: {RefCount}", 
                    objectType, context, refCount);
                
                return true;
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Failed to release COM object {ObjectType} in {Context}", 
                    objectType, context);
                return false;
            }
        }

        /// <summary>
        /// Tracks the creation of a COM object
        /// </summary>
        /// <param name="objectType">Type of object created</param>
        /// <param name="context">Context information</param>
        public static void TrackComObjectCreation(string objectType, string context = "")
        {
            Interlocked.Increment(ref _totalCreated);
            
            var key = $"{objectType}:{context}";
            _stats.AddOrUpdate(key, 
                new ComObjectStats { ObjectType = objectType, Context = context, Created = 1 },
                (k, existing) => 
                {
                    existing.Created++;
                    return existing;
                });

            _logger.Debug("Created COM object {ObjectType} in {Context}, Total Created: {TotalCreated}", 
                objectType, context, _totalCreated);
        }

        /// <summary>
        /// Tracks the disposal of a COM object
        /// </summary>
        /// <param name="objectType">Type of object disposed</param>
        /// <param name="context">Context information</param>
        public static void TrackComObjectDisposal(string objectType, string context = "")
        {
            Interlocked.Increment(ref _totalReleased);
            
            var key = $"{objectType}:{context}";
            _stats.AddOrUpdate(key,
                new ComObjectStats { ObjectType = objectType, Context = context, Released = 1 },
                (k, existing) =>
                {
                    existing.Released++;
                    return existing;
                });

            _logger.Debug("Released COM object {ObjectType} in {Context}, Total Released: {TotalReleased}", 
                objectType, context, _totalReleased);
        }

        /// <summary>
        /// Gets current COM object statistics
        /// </summary>
        /// <returns>Summary of COM object creation/disposal counts</returns>
        public static ComObjectSummary GetComObjectSummary()
        {
            var summary = new ComObjectSummary
            {
                TotalCreated = _totalCreated,
                TotalReleased = _totalReleased,
                NetObjects = _totalCreated - _totalReleased,
                ObjectStats = new Dictionary<string, ComObjectStats>()
            };

            foreach (var kvp in _stats)
            {
                summary.ObjectStats[kvp.Key] = new ComObjectStats
                {
                    ObjectType = kvp.Value.ObjectType,
                    Context = kvp.Value.Context,
                    Created = kvp.Value.Created,
                    Released = kvp.Value.Released
                };
            }

            return summary;
        }

        /// <summary>
        /// Gets simplified COM object statistics for health monitoring
        /// </summary>
        /// <returns>Object with NetObjects count</returns>
        public static ComObjectSummary GetComObjectStats()
        {
            return GetComObjectSummary();
        }
        
        /// <summary>
        /// Logs current COM object statistics
        /// </summary>
        public static void LogComObjectStats()
        {
            var summary = GetComObjectSummary();
            
            _logger.Information("COM Object Statistics - Created: {Created}, Released: {Released}, Net: {Net}", 
                summary.TotalCreated, summary.TotalReleased, summary.NetObjects);

            if (summary.NetObjects > 0)
            {
                _logger.Warning("Potential COM object leak detected - {Net} objects not released", 
                    summary.NetObjects);
            }

            // Log details for each object type
            foreach (var kvp in summary.ObjectStats)
            {
                var stats = kvp.Value;
                var net = stats.Created - stats.Released;
                
                if (net != 0)
                {
                    var level = net > 0 ? "Warning" : "Information";
                    _logger.Write(net > 0 ? Serilog.Events.LogEventLevel.Warning : Serilog.Events.LogEventLevel.Information,
                        "COM Object {ObjectType} ({Context}): Created {Created}, Released {Released}, Net {Net}",
                        stats.ObjectType, stats.Context, stats.Created, stats.Released, net);
                }
            }
        }

        /// <summary>
        /// Resets COM object tracking statistics
        /// </summary>
        public static void ResetStats()
        {
            _stats.Clear();
            Interlocked.Exchange(ref _totalCreated, 0);
            Interlocked.Exchange(ref _totalReleased, 0);
            
            _logger.Information("COM object tracking statistics reset");
        }

        /// <summary>
        /// Performs aggressive garbage collection to clean up COM objects
        /// </summary>
        public static void ForceComCleanup(string context = "")
        {
            _logger.Debug("Forcing COM cleanup in {Context}", context);
            
            // Triple garbage collection pattern for COM cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            
            _logger.Debug("COM cleanup completed in {Context}", context);
        }

        /// <summary>
        /// Checks if COM object counts are balanced (created = released)
        /// </summary>
        /// <returns>True if balanced, false if potential leaks detected</returns>
        public static bool AreComObjectsBalanced()
        {
            var summary = GetComObjectSummary();
            return summary.NetObjects == 0;
        }

        /// <summary>
        /// Gets the count of potentially leaked COM objects
        /// </summary>
        /// <returns>Number of objects created but not released</returns>
        public static long GetLeakedObjectCount()
        {
            var summary = GetComObjectSummary();
            return Math.Max(0, summary.NetObjects);
        }
    }

    /// <summary>
    /// Statistics for a specific type of COM object
    /// </summary>
    public class ComObjectStats
    {
        public string ObjectType { get; set; } = "";
        public string Context { get; set; } = "";
        public long Created { get; set; }
        public long Released { get; set; }
        public long Net => Created - Released;
    }

    /// <summary>
    /// Summary of all COM object statistics
    /// </summary>
    public class ComObjectSummary
    {
        public long TotalCreated { get; set; }
        public long TotalReleased { get; set; }
        public long NetObjects { get; set; }
        public Dictionary<string, ComObjectStats> ObjectStats { get; set; } = new();
    }
} 