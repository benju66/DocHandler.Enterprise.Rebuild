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
            services.AddSingleton<MemoryLeakDetector>();
            
            // Data Services (Singleton - shared data access)
            services.AddSingleton<ScopeOfWorkService>();
            services.AddSingleton<CompanyNameService>();
            
            // PDF Operations (Transient - stateless operations)
            services.AddTransient<PdfOperationsService>();
            
            return services;
        }
    }
} 