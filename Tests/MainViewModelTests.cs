using Microsoft.Extensions.DependencyInjection;
using DocHandler.Services;
using DocHandler.ViewModels;
using Xunit;
using Moq;
using System;
using Serilog;

namespace DocHandler.Tests
{
    public class MainViewModelTests : IDisposable
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly Mock<IConfigurationService> _mockConfigService;
        private readonly Mock<IFileProcessingService> _mockFileProcessingService;
        private readonly Mock<ICompanyNameService> _mockCompanyNameService;

        public MainViewModelTests()
        {
            // Setup mocks
            _mockConfigService = new Mock<IConfigurationService>();
            _mockFileProcessingService = new Mock<IFileProcessingService>();
            _mockCompanyNameService = new Mock<ICompanyNameService>();

            // Setup mock configuration
            var mockConfig = new AppConfiguration
            {
                MemoryUsageLimitMB = 512,
                DocFileSizeLimitMB = 10
            };
            _mockConfigService.Setup(x => x.Config).Returns(mockConfig);

            // Setup logger
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .CreateLogger();

            // Setup DI container with mocks
            var services = new ServiceCollection();
            
            // Register mocks
            services.AddSingleton(_mockConfigService.Object);
            services.AddSingleton(_mockFileProcessingService.Object);
            services.AddSingleton(_mockCompanyNameService.Object);
            
            // Register other required services with mocks or real implementations
            services.AddSingleton<IProcessManager, ProcessManager>();
            services.AddSingleton<IPerformanceMonitor, PerformanceMonitor>();
            services.AddSingleton<IPdfCacheService, PdfCacheService>();
            services.AddSingleton<IScopeOfWorkService, ScopeOfWorkService>();
            services.AddSingleton<ITelemetryService, TelemetryService>();
            services.AddTransient<IPdfOperationsService, PdfOperationsService>();
            
            // Mock Office services to avoid COM dependencies in tests
            services.AddScoped<ISessionAwareOfficeService>(_ => Mock.Of<ISessionAwareOfficeService>());
            services.AddScoped<ISessionAwareExcelService>(_ => Mock.Of<ISessionAwareExcelService>());
            services.AddScoped<IOfficeConversionService>(_ => Mock.Of<IOfficeConversionService>());
            services.AddScoped<ISaveQuotesQueueService>(_ => Mock.Of<ISaveQuotesQueueService>());

            _serviceProvider = services.BuildServiceProvider();
        }

        [Fact]
        public void MainViewModel_Constructor_ShouldAcceptServiceProvider()
        {
            // Arrange & Act
            var viewModel = new MainViewModel(_serviceProvider);

            // Assert
            Assert.NotNull(viewModel);
            Assert.NotNull(viewModel.ConfigService);
        }

        [Fact]
        public void MainViewModel_Constructor_ShouldInjectAllRequiredServices()
        {
            // Arrange & Act
            var viewModel = new MainViewModel(_serviceProvider);

            // Assert
            Assert.NotNull(viewModel.ConfigService);
            Assert.Same(_mockConfigService.Object, viewModel.ConfigService);
        }

        [Fact]
        public void MainViewModel_LegacyConstructor_ShouldCreateServiceProvider()
        {
            // Arrange & Act
            var viewModel = new MainViewModel();

            // Assert
            Assert.NotNull(viewModel);
            Assert.NotNull(viewModel.ConfigService);
        }

        [Fact]
        public void MainViewModel_Constructor_WithNullServiceProvider_ShouldThrowArgumentNullException()
        {
            // Arrange, Act & Assert
            Assert.Throws<ArgumentNullException>(() => new MainViewModel(null!));
        }

        [Fact]
        public void MainViewModel_ConfigService_ShouldReturnInjectedService()
        {
            // Arrange
            var viewModel = new MainViewModel(_serviceProvider);

            // Act
            var configService = viewModel.ConfigService;

            // Assert
            Assert.NotNull(configService);
            Assert.Same(_mockConfigService.Object, configService);
        }

        [Fact]
        public void MainViewModel_PropertyChanges_ShouldCallConfigurationService()
        {
            // Arrange
            var viewModel = new MainViewModel(_serviceProvider);
            _mockConfigService.Setup(x => x.Config).Returns(new AppConfiguration { DocFileSizeLimitMB = 5 });

            // Act
            viewModel.DocFileSizeLimitMB = 20;

            // Assert
            _mockConfigService.Verify(x => x.Config, Times.AtLeastOnce);
        }

        [Fact]
        public void MainViewModel_Dispose_ShouldDisposeResources()
        {
            // Arrange
            var viewModel = new MainViewModel(_serviceProvider);

            // Act & Assert - Should not throw
            viewModel.Dispose();
        }

        [Fact]
        public void MainViewModel_MultipleInstances_ShouldUseSameSingletonServices()
        {
            // Arrange & Act
            var viewModel1 = new MainViewModel(_serviceProvider);
            var viewModel2 = new MainViewModel(_serviceProvider);

            // Assert
            Assert.Same(viewModel1.ConfigService, viewModel2.ConfigService);
        }

        public void Dispose()
        {
            _serviceProvider?.Dispose();
            Log.CloseAndFlush();
        }
    }

    // Mock AppConfiguration for testing
    public class AppConfiguration
    {
        public int MemoryUsageLimitMB { get; set; } = 512;
        public int DocFileSizeLimitMB { get; set; } = 10;
        public string SaveLocation { get; set; } = "";
        public bool PreferExactMatches { get; set; } = true;
        public bool SaveQuotesMode { get; set; } = false;
        public bool ShowQueueWindow { get; set; } = false;
    }
}