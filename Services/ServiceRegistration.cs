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
            services.AddSingleton<IConfigurationService, ConfigurationService>();
            services.AddSingleton<IProcessManager, ProcessManager>();
            services.AddSingleton<IPerformanceMonitor, PerformanceMonitor>();
            services.AddSingleton<ITelemetryService, TelemetryService>();
            services.AddSingleton<IPdfCacheService, PdfCacheService>();
            
            // Data Services (Singleton - shared data access)
            services.AddSingleton<IScopeOfWorkService, ScopeOfWorkService>();
            services.AddSingleton<ICompanyNameService, CompanyNameService>();
            
            // PDF Operations (Transient - stateless operations)
            services.AddTransient<IPdfOperationsService, PdfOperationsService>();
            
            // Office Services (Scoped - session-aware, but need careful lifetime management)
            services.AddScoped<ISessionAwareOfficeService, SessionAwareOfficeService>();
            services.AddScoped<ISessionAwareExcelService, SessionAwareExcelService>();
            services.AddScoped<IOfficeConversionService, OfficeConversionService>();
            
            // Processing Services (Scoped - per operation lifecycle)
            services.AddScoped<IFileProcessingService, OptimizedFileProcessingService>();
            services.AddScoped<ISaveQuotesQueueService, SaveQuotesQueueService>();
            
            // Health and Monitoring Services (Singleton - application-wide monitoring)
            services.AddSingleton<OfficeHealthMonitor>();
            services.AddSingleton<ApplicationHealthChecker>();
            
            // Utility Services (Transient - lightweight, stateless)
            services.AddTransient<CircuitBreaker>();
            services.AddTransient<ConversionCircuitBreaker>();
            services.AddTransient<ComHelper>();
            services.AddTransient<StaThreadPool>();
            
            return services;
        }
    }
} 