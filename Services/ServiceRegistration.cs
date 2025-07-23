using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using Serilog;
using DocHandler.ViewModels;
using DocHandler.Services.Pipeline;
using DocHandler.Services.Pipeline.SaveQuotes;
using DocHandler.Services.Configuration;

namespace DocHandler.Services
{
    public static class ServiceRegistration
    {
        public static IServiceCollection RegisterServices(this IServiceCollection services)
        {
            return RegisterServices(services, null);
        }

        public static IServiceCollection RegisterServices(this IServiceCollection services, IHierarchicalConfigurationService? configService)
        {
            // Core Infrastructure Services (Singleton - long-running resources)
            services.AddSingleton<IConfigurationService, ConfigurationService>();
            services.AddSingleton<IHierarchicalConfigurationService, HierarchicalConfigurationService>();
            services.AddSingleton<IConfigurationChangeNotificationService, ConfigurationChangeNotificationService>();
            services.AddTransient<IConfigurationExportImportService, ConfigurationExportImportService>();
            services.AddSingleton<IProcessManager, ProcessManager>();
            services.AddSingleton<PerformanceMonitor>(provider => 
                new PerformanceMonitor(
                    provider.GetService<IHierarchicalConfigurationService>(),
                    provider.GetService<IConfigurationChangeNotificationService>()));
            // Application Insights and Telemetry
            services.AddApplicationInsightsTelemetryWorkerService(options =>
            {
                if (configService?.Config?.Telemetry != null)
                {
                    var telemetryConfig = configService.Config.Telemetry;
                    
                    if (telemetryConfig.EnableApplicationInsights)
                    {
                        if (!string.IsNullOrEmpty(telemetryConfig.ApplicationInsightsConnectionString))
                        {
                            options.ConnectionString = telemetryConfig.ApplicationInsightsConnectionString;
                        }
                        else if (!string.IsNullOrEmpty(telemetryConfig.ApplicationInsightsInstrumentationKey))
                        {
                            options.InstrumentationKey = telemetryConfig.ApplicationInsightsInstrumentationKey;
                        }
                    }
                }
            });
            
            // Configure Application Insights telemetry configuration
            services.Configure<TelemetryConfiguration>(telemetryConfig =>
            {
                if (configService?.Config?.Telemetry != null)
                {
                    var config = configService.Config.Telemetry;
                    
                    if (!config.EnableApplicationInsights)
                    {
                        telemetryConfig.DisableTelemetry = true;
                    }
                    
                    // Set sampling percentage
                    if (config.SamplingPercentage > 0 && config.SamplingPercentage < 100)
                    {
                        telemetryConfig.DefaultTelemetrySink.TelemetryProcessorChainBuilder
                            .UseSampling(config.SamplingPercentage)
                            .Build();
                    }
                }
            });
            services.AddSingleton<TelemetryService>(provider =>
            {
                var performanceMonitor = provider.GetService<PerformanceMonitor>();
                var applicationInsightsClient = provider.GetService<TelemetryClient>();
                var configService = provider.GetService<IHierarchicalConfigurationService>();
                
                return new TelemetryService(performanceMonitor, applicationInsightsClient, configService);
            });
            
            // WPF-specific telemetry helper
            services.AddSingleton<WpfTelemetryHelper>(provider =>
            {
                var telemetryService = provider.GetService<TelemetryService>();
                var performanceMonitor = provider.GetService<PerformanceMonitor>();
                
                return new WpfTelemetryHelper(telemetryService, performanceMonitor);
            });
            
            // Activity tracing for correlation IDs and workflow tracking
            services.AddSingleton<ActivityTracingService>(provider =>
            {
                var telemetryService = provider.GetService<TelemetryService>();
                var configService = provider.GetService<IHierarchicalConfigurationService>();
                var applicationInsightsClient = provider.GetService<TelemetryClient>();
                
                return new ActivityTracingService(telemetryService, configService, applicationInsightsClient);
            });
            
            services.AddSingleton<PdfCacheService>();
            
            // Data Services (Singleton - shared data access)
            services.AddSingleton<IScopeOfWorkService, ScopeOfWorkService>();
            services.AddSingleton<ICompanyNameService, CompanyNameService>();
            
            // File Processing Services (Transient - stateful per operation)
            services.AddTransient<IOptimizedFileProcessingService, OptimizedFileProcessingService>();
            services.AddTransient<ISaveQuotesQueueService, SaveQuotesQueueService>();
            
            // Office Services (Transient - COM object management)
            services.AddTransient<IOfficeConversionService, OfficeConversionService>();
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
            
            // Business Services (Phase 2)
            services.AddSingleton<IEnhancedFileValidationService, FileValidationService>();
            services.AddSingleton<IFileValidationService>(provider => (IFileValidationService)provider.GetRequiredService<IEnhancedFileValidationService>());
            services.AddSingleton<IEnhancedCompanyDetectionService, CompanyDetectionService>();
            services.AddSingleton<ICompanyDetectionService>(provider => (ICompanyDetectionService)provider.GetRequiredService<IEnhancedCompanyDetectionService>());
            services.AddSingleton<IDocumentWorkflowService, DocumentWorkflowService>();
            services.AddSingleton<IScopeManagementService, ScopeManagementService>();
            services.AddSingleton<IUIStateService, UIStateService>();

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