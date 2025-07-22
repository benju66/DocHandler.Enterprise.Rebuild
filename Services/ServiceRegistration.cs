using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;

namespace DocHandler.Services
{
    public static class ServiceRegistration
    {
        public static IServiceCollection RegisterServices(this IServiceCollection services)
        {
            // Core Infrastructure Services (Singleton - long-running resources)
            // Register both concrete and interface types for services that implement interfaces
            services.AddSingleton<IConfigurationService, ConfigurationService>();
            services.AddSingleton<IProcessManager, ProcessManager>();
            services.AddSingleton<PerformanceMonitor>();
            services.AddSingleton<TelemetryService>();
            services.AddSingleton<PdfCacheService>();
            
            // Data Services (Singleton - shared data access)
            services.AddSingleton<ScopeOfWorkService>();
            services.AddSingleton<CompanyNameService>();
            
            // PDF Operations (Transient - stateless operations)
            services.AddTransient<PdfOperationsService>();
            
            // Mode Infrastructure (Singleton - mode management)
            services.AddModeRegistry();
            services.AddSingleton<IModeManager, ModeManager>();
            
            // Error Recovery Service (Singleton - error handling)
            services.AddSingleton<ErrorRecoveryService>();
            
            // Office Health Monitor (Singleton - health checking)
            services.AddSingleton<OfficeHealthMonitor>();
            
            return services;
        }

        /// <summary>
        /// Register processing modes
        /// </summary>
        public static IServiceCollection RegisterProcessingModes(this IServiceCollection services)
        {
            // Register SaveQuotes mode
            services.RegisterProcessingMode<Modes.SaveQuotesMode>();
            
            // Future modes will be registered here
            
            return services;
        }

        /// <summary>
        /// Configure mode-specific services
        /// </summary>
        public static IServiceCollection ConfigureModeServices(this IServiceCollection services)
        {
            // Mode-specific service configurations can be added here
            // For example, different validation rules, different processors, etc.
            
            return services;
        }
    }
} 