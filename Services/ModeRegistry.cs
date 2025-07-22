using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Registry for managing processing modes
    /// </summary>
    public interface IModeRegistry
    {
        /// <summary>
        /// Register a mode type
        /// </summary>
        void RegisterMode<TMode>() where TMode : class, IProcessingMode;
        
        /// <summary>
        /// Register a mode instance
        /// </summary>
        void RegisterMode(string modeName, Func<IServiceProvider, IProcessingMode> factory);
        
        /// <summary>
        /// Get all registered mode descriptors
        /// </summary>
        IEnumerable<ModeDescriptor> GetAvailableModes();
        
        /// <summary>
        /// Check if a mode is registered
        /// </summary>
        bool IsRegistered(string modeName);
        
        /// <summary>
        /// Create a mode instance
        /// </summary>
        Task<IProcessingMode> CreateModeAsync(string modeName, IServiceProvider serviceProvider);
        
        /// <summary>
        /// Get mode descriptor by name
        /// </summary>
        ModeDescriptor? GetModeDescriptor(string modeName);
    }

    /// <summary>
    /// Descriptor for a processing mode
    /// </summary>
    public class ModeDescriptor
    {
        public string ModeName { get; set; } = string.Empty;
        public string DisplayName { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public Version Version { get; set; } = new Version(1, 0, 0);
        public Type ModeType { get; set; } = typeof(object);
        public bool IsAvailable { get; set; } = true;
        public Func<IServiceProvider, IProcessingMode>? Factory { get; set; }
    }

    /// <summary>
    /// Implementation of mode registry
    /// </summary>
    public class ModeRegistry : IModeRegistry
    {
        private readonly ILogger _logger;
        private readonly ConcurrentDictionary<string, ModeDescriptor> _modes;

        public ModeRegistry()
        {
            _logger = Log.ForContext<ModeRegistry>();
            _modes = new ConcurrentDictionary<string, ModeDescriptor>();
        }

        public void RegisterMode<TMode>() where TMode : class, IProcessingMode
        {
            var modeType = typeof(TMode);
            
            // Create a temporary instance to get metadata
            try
            {
                var tempInstance = ActivatorUtilities.CreateInstance<TMode>(CreateMinimalServiceProvider());
                
                var descriptor = new ModeDescriptor
                {
                    ModeName = tempInstance.ModeName,
                    DisplayName = tempInstance.DisplayName,
                    Description = tempInstance.Description,
                    Version = tempInstance.Version,
                    ModeType = modeType,
                    IsAvailable = tempInstance.IsAvailable,
                    Factory = serviceProvider => ActivatorUtilities.CreateInstance<TMode>(serviceProvider)
                };

                _modes.TryAdd(descriptor.ModeName, descriptor);
                _logger.Information("Registered mode: {ModeName} ({DisplayName}) v{Version}", 
                    descriptor.ModeName, descriptor.DisplayName, descriptor.Version);
                
                tempInstance.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to register mode {ModeType}", modeType.Name);
                throw;
            }
        }

        public void RegisterMode(string modeName, Func<IServiceProvider, IProcessingMode> factory)
        {
            if (string.IsNullOrWhiteSpace(modeName))
                throw new ArgumentException("Mode name cannot be null or empty", nameof(modeName));
            
            if (factory == null)
                throw new ArgumentNullException(nameof(factory));

            try
            {
                // Create a temporary instance to get metadata
                var tempInstance = factory(CreateMinimalServiceProvider());
                
                var descriptor = new ModeDescriptor
                {
                    ModeName = modeName,
                    DisplayName = tempInstance.DisplayName,
                    Description = tempInstance.Description,
                    Version = tempInstance.Version,
                    ModeType = tempInstance.GetType(),
                    IsAvailable = tempInstance.IsAvailable,
                    Factory = factory
                };

                _modes.TryAdd(modeName, descriptor);
                _logger.Information("Registered mode: {ModeName} ({DisplayName}) v{Version}", 
                    descriptor.ModeName, descriptor.DisplayName, descriptor.Version);
                
                tempInstance.Dispose();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to register mode {ModeName}", modeName);
                throw;
            }
        }

        public IEnumerable<ModeDescriptor> GetAvailableModes()
        {
            return _modes.Values.Where(d => d.IsAvailable).ToList();
        }

        public bool IsRegistered(string modeName)
        {
            return !string.IsNullOrWhiteSpace(modeName) && _modes.ContainsKey(modeName);
        }

        public async Task<IProcessingMode> CreateModeAsync(string modeName, IServiceProvider serviceProvider)
        {
            if (string.IsNullOrWhiteSpace(modeName))
                throw new ArgumentException("Mode name cannot be null or empty", nameof(modeName));
            
            if (serviceProvider == null)
                throw new ArgumentNullException(nameof(serviceProvider));

            if (!_modes.TryGetValue(modeName, out var descriptor))
                throw new InvalidOperationException($"Mode '{modeName}' is not registered");

            if (descriptor.Factory == null)
                throw new InvalidOperationException($"Mode '{modeName}' has no factory");

            try
            {
                _logger.Debug("Creating mode instance: {ModeName}", modeName);
                var mode = descriptor.Factory(serviceProvider);
                
                // Initialize the mode
                var context = new ModeContext(serviceProvider);
                await mode.InitializeAsync(context);
                
                _logger.Information("Created and initialized mode: {ModeName}", modeName);
                return mode;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Failed to create mode: {ModeName}", modeName);
                throw;
            }
        }

        public ModeDescriptor? GetModeDescriptor(string modeName)
        {
            return string.IsNullOrWhiteSpace(modeName) ? null :
                _modes.TryGetValue(modeName, out var descriptor) ? descriptor : null;
        }

        /// <summary>
        /// Create a minimal service provider for temporary mode creation
        /// </summary>
        private static IServiceProvider CreateMinimalServiceProvider()
        {
            var services = new ServiceCollection();
            
            // Add minimal required services
            services.AddLogging();
            services.AddSingleton<ILogger>(Log.Logger);
            
            return services.BuildServiceProvider();
        }
    }

    /// <summary>
    /// Service collection extensions for mode registration
    /// </summary>
    public static class ModeRegistryExtensions
    {
        /// <summary>
        /// Add mode registry to service collection
        /// </summary>
        public static IServiceCollection AddModeRegistry(this IServiceCollection services)
        {
            services.AddSingleton<IModeRegistry, ModeRegistry>();
            return services;
        }
        
        /// <summary>
        /// Register a mode in the service collection
        /// </summary>
        public static IServiceCollection RegisterProcessingMode<TMode>(this IServiceCollection services) 
            where TMode : class, IProcessingMode
        {
            services.AddTransient<TMode>();
            return services;
        }
    }
} 