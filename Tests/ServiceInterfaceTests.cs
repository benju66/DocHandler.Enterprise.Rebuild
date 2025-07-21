using DocHandler.Services;
using Xunit;
using System;
using System.Reflection;
using System.Linq;
using Microsoft.Extensions.DependencyInjection;

namespace DocHandler.Tests
{
    public class ServiceInterfaceTests
    {
        [Fact]
        public void ConfigurationService_ShouldImplementIConfigurationService()
        {
            // Arrange
            var serviceType = typeof(ConfigurationService);
            var interfaceType = typeof(IConfigurationService);

            // Act & Assert
            Assert.True(interfaceType.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {interfaceType.Name}");
        }

        [Fact]
        public void PdfCacheService_ShouldImplementIPdfCacheService()
        {
            // Arrange
            var serviceType = typeof(PdfCacheService);
            var interfaceType = typeof(IPdfCacheService);

            // Act & Assert
            Assert.True(interfaceType.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {interfaceType.Name}");
        }

        [Fact]
        public void SessionAwareOfficeService_ShouldImplementISessionAwareOfficeService()
        {
            // Arrange
            var serviceType = typeof(SessionAwareOfficeService);
            var interfaceType = typeof(ISessionAwareOfficeService);

            // Act & Assert
            Assert.True(interfaceType.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {interfaceType.Name}");
        }

        [Fact]
        public void CompanyNameService_ShouldImplementICompanyNameService()
        {
            // Arrange
            var serviceType = typeof(CompanyNameService);
            var interfaceType = typeof(ICompanyNameService);

            // Act & Assert
            Assert.True(interfaceType.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {interfaceType.Name}");
        }

        [Fact]
        public void OptimizedFileProcessingService_ShouldImplementIFileProcessingService()
        {
            // Arrange
            var serviceType = typeof(OptimizedFileProcessingService);
            var interfaceType = typeof(IFileProcessingService);

            // Act & Assert
            Assert.True(interfaceType.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {interfaceType.Name}");
        }

        [Fact]
        public void SaveQuotesQueueService_ShouldImplementISaveQuotesQueueService()
        {
            // Arrange
            var serviceType = typeof(SaveQuotesQueueService);
            var interfaceType = typeof(ISaveQuotesQueueService);

            // Act & Assert
            Assert.True(interfaceType.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {interfaceType.Name}");
        }

        [Fact]
        public void AllServiceInterfaces_ShouldInheritFromIDisposable()
        {
            // Arrange
            var interfaceTypes = new[]
            {
                typeof(IConfigurationService),
                typeof(IPdfCacheService),
                typeof(ISessionAwareOfficeService),
                typeof(ISessionAwareExcelService),
                typeof(ICompanyNameService),
                typeof(IFileProcessingService),
                typeof(IOfficeConversionService),
                typeof(ISaveQuotesQueueService),
                typeof(IPerformanceMonitor),
                typeof(ITelemetryService)
            };

            // Act & Assert
            foreach (var interfaceType in interfaceTypes)
            {
                Assert.True(typeof(IDisposable).IsAssignableFrom(interfaceType),
                    $"{interfaceType.Name} should inherit from IDisposable");
            }
        }

        [Fact]
        public void IConfigurationService_ShouldHaveRequiredMethods()
        {
            // Arrange
            var interfaceType = typeof(IConfigurationService);

            // Act
            var methods = interfaceType.GetMethods();
            var properties = interfaceType.GetProperties();

            // Assert
            Assert.Contains(methods, m => m.Name == "SaveConfigurationAsync");
            Assert.Contains(methods, m => m.Name == "UpdateConfiguration");
            Assert.Contains(properties, p => p.Name == "Config");
        }

        [Fact]
        public void IFileProcessingService_ShouldHaveRequiredMethods()
        {
            // Arrange
            var interfaceType = typeof(IFileProcessingService);

            // Act
            var methods = interfaceType.GetMethods();

            // Assert
            Assert.Contains(methods, m => m.Name == "IsFileSupported");
            Assert.Contains(methods, m => m.Name == "ProcessFilesAsync");
            Assert.Contains(methods, m => m.Name == "ConvertToPdfAsync");
        }

        [Fact]
        public void ISaveQuotesQueueService_ShouldHaveRequiredMethods()
        {
            // Arrange
            var interfaceType = typeof(ISaveQuotesQueueService);

            // Act
            var methods = interfaceType.GetMethods();
            var properties = interfaceType.GetProperties();
            var events = interfaceType.GetEvents();

            // Assert
            Assert.Contains(methods, m => m.Name == "AddToQueue");
            Assert.Contains(methods, m => m.Name == "StartProcessingAsync");
            Assert.Contains(methods, m => m.Name == "StopProcessingAsync");
            Assert.Contains(methods, m => m.Name == "ClearQueueAsync");
            
            Assert.Contains(properties, p => p.Name == "IsProcessing");
            Assert.Contains(properties, p => p.Name == "TotalCount");
            Assert.Contains(properties, p => p.Name == "ProcessedCount");
            Assert.Contains(properties, p => p.Name == "FailedCount");
            
            Assert.Contains(events, e => e.Name == "ProgressChanged");
            Assert.Contains(events, e => e.Name == "ItemCompleted");
            Assert.Contains(events, e => e.Name == "QueueEmpty");
        }

        [Theory]
        [InlineData(typeof(ConfigurationService), typeof(IConfigurationService))]
        [InlineData(typeof(PdfCacheService), typeof(IPdfCacheService))]
        [InlineData(typeof(SessionAwareOfficeService), typeof(ISessionAwareOfficeService))]
        [InlineData(typeof(SessionAwareExcelService), typeof(ISessionAwareExcelService))]
        [InlineData(typeof(CompanyNameService), typeof(ICompanyNameService))]
        [InlineData(typeof(ScopeOfWorkService), typeof(IScopeOfWorkService))]
        [InlineData(typeof(OptimizedFileProcessingService), typeof(IFileProcessingService))]
        [InlineData(typeof(OfficeConversionService), typeof(IOfficeConversionService))]
        [InlineData(typeof(SaveQuotesQueueService), typeof(ISaveQuotesQueueService))]
        [InlineData(typeof(PdfOperationsService), typeof(IPdfOperationsService))]
        [InlineData(typeof(TelemetryService), typeof(ITelemetryService))]
        [InlineData(typeof(PerformanceMonitor), typeof(IPerformanceMonitor))]
        public void Service_ShouldImplementCorrectInterface(Type serviceType, Type expectedInterface)
        {
            // Act & Assert
            Assert.True(expectedInterface.IsAssignableFrom(serviceType),
                $"{serviceType.Name} should implement {expectedInterface.Name}");
        }

        [Fact]
        public void AllRegisteredServices_ShouldBeResolvableFromDI()
        {
            // Arrange
            var services = new ServiceCollection();
            services.RegisterServices();
            var serviceProvider = services.BuildServiceProvider();

            var interfaceTypes = new[]
            {
                typeof(IConfigurationService),
                typeof(IProcessManager),
                typeof(IPerformanceMonitor),
                typeof(ITelemetryService),
                typeof(IPdfCacheService),
                typeof(IScopeOfWorkService),
                typeof(ICompanyNameService),
                typeof(IPdfOperationsService)
            };

            // Act & Assert
            foreach (var interfaceType in interfaceTypes)
            {
                var service = serviceProvider.GetService(interfaceType);
                Assert.NotNull(service);
                Assert.True(interfaceType.IsAssignableFrom(service.GetType()),
                    $"Service resolved for {interfaceType.Name} should implement the interface");
            }

            serviceProvider.Dispose();
        }
    }
}