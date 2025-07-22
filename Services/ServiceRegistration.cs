using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using DocHandler.ViewModels;
using DocHandler.Services.Pipeline;
using DocHandler.Services.Pipeline.SaveQuotes;

namespace DocHandler.Services
{
    public static class ServiceRegistration
    {
        public static IServiceCollection RegisterServices(this IServiceCollection services)
        {
            // Core Infrastructure Services (Singleton - long-running resources)
            services.AddSingleton<IConfigurationService, ConfigurationService>();
            services.AddSingleton<IProcessManager, ProcessManager>();
            services.AddSingleton<PerformanceMonitor>();
            services.AddSingleton<TelemetryService>();
            services.AddSingleton<PdfCacheService>();
            
            // Data Services (Singleton - shared data access)
            services.AddSingleton<IScopeOfWorkService, ScopeOfWorkService>();
            services.AddSingleton<ICompanyNameService, CompanyNameService>();
            
            // File Processing Services (Transient - stateful per operation)
            services.AddTransient<IOptimizedFileProcessingService, OptimizedFileProcessingService>();
            services.AddTransient<ISaveQuotesQueueService, SaveQuotesQueueService>();
            
            // Office Services (Transient - COM object management)
            services.AddTransient<ISessionAwareOfficeService, SessionAwareOfficeService>();
            services.AddTransient<ISessionAwareExcelService, SessionAwareExcelService>();
            
            // PDF Operations (Transient - stateless operations)
            services.AddTransient<PdfOperationsService>();
            
            // Mode Infrastructure (Singleton - mode management)
            services.AddModeRegistry();
            services.AddSingleton<IModeManager, ModeManager>();
            
            // Error Recovery Service (Singleton - error handling)
            services.AddSingleton<ErrorRecoveryService>();
            
            // Office Health Monitor (Singleton - health checking)
            services.AddSingleton<OfficeHealthMonitor>();
            
            // Business Logic Services (Transient - stateful per operation)
            services.AddTransient<IFileValidationService, FileValidationService>();
            services.AddTransient<ICompanyDetectionService, CompanyDetectionService>();
            services.AddTransient<IScopeManagementService, ScopeManagementService>();

            // Mode UI Framework Services (Phase 2 Milestone 2 - Day 4)
            services.AddSingleton<IAdvancedModeUIProvider, ModeUIProvider>();
            services.AddSingleton<IDynamicMenuBuilder, DynamicMenuBuilder>();
            services.AddSingleton<IAdvancedModeUIManager, ModeUIManager>();

            // Processing Pipeline Framework (Phase 2 Milestone 3)
            services.AddTransient<IPipelineBuilder, PipelineBuilder>();
            services.AddTransient<IProcessingPipeline, ProcessingPipeline>();

            // SaveQuotes Pipeline Stages (Phase 2 Milestone 3)
            services.AddTransient<SaveQuotesValidator>();
            services.AddTransient<SaveQuotesPreProcessor>();
            services.AddTransient<SaveQuotesConverter>();
            services.AddTransient<SaveQuotesPostProcessor>();
            services.AddTransient<SaveQuotesOutputGenerator>();
            
            // ViewModels (Transient - new instance per resolution)
            services.AddTransient<ViewModels.MainViewModel>();
            
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