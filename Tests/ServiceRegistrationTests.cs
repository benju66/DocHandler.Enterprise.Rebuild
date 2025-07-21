using Microsoft.Extensions.DependencyInjection;
using DocHandler.Services;
using Xunit;
using System;
using System.Collections.Generic;

namespace DocHandler.Tests
{
    public class ServiceRegistrationTests
    {
        private readonly IServiceProvider _serviceProvider;

        public ServiceRegistrationTests()
        {
            var services = new ServiceCollection();
            services.RegisterServices();
            _serviceProvider = services.BuildServiceProvider();
        }

        [Fact]
        public void RegisterServices_ShouldRegisterAllCoreServices()
        {
            // Arrange & Act - services are registered in constructor

            // Assert - Core Infrastructure Services (Singleton)
            Assert.NotNull(_serviceProvider.GetService<IConfigurationService>());
            Assert.NotNull(_serviceProvider.GetService<IProcessManager>());
            Assert.NotNull(_serviceProvider.GetService<IPerformanceMonitor>());
            Assert.NotNull(_serviceProvider.GetService<ITelemetryService>());
            Assert.NotNull(_serviceProvider.GetService<IPdfCacheService>());
        }

        [Fact]
        public void RegisterServices_ShouldRegisterDataServices()
        {
            // Assert - Data Services (Singleton)
            Assert.NotNull(_serviceProvider.GetService<IScopeOfWorkService>());
            Assert.NotNull(_serviceProvider.GetService<ICompanyNameService>());
        }

        [Fact]
        public void RegisterServices_ShouldRegisterProcessingServices()
        {
            // Assert - Processing Services (Scoped)
            Assert.NotNull(_serviceProvider.GetService<IFileProcessingService>());
            Assert.NotNull(_serviceProvider.GetService<ISaveQuotesQueueService>());
        }

        [Fact]
        public void RegisterServices_ShouldRegisterOfficeServices()
        {
            // Assert - Office Services (Scoped)
            Assert.NotNull(_serviceProvider.GetService<ISessionAwareOfficeService>());
            Assert.NotNull(_serviceProvider.GetService<ISessionAwareExcelService>());
            Assert.NotNull(_serviceProvider.GetService<IOfficeConversionService>());
        }

        [Fact]
        public void RegisterServices_ShouldRegisterUtilityServices()
        {
            // Assert - PDF Operations (Transient)
            Assert.NotNull(_serviceProvider.GetService<IPdfOperationsService>());
            
            // Assert - Utility Services (Transient)
            Assert.NotNull(_serviceProvider.GetService<CircuitBreaker>());
            Assert.NotNull(_serviceProvider.GetService<ConversionCircuitBreaker>());
            Assert.NotNull(_serviceProvider.GetService<ComHelper>());
            Assert.NotNull(_serviceProvider.GetService<StaThreadPool>());
        }

        [Fact]
        public void RegisterServices_SingletonServices_ShouldReturnSameInstance()
        {
            // Arrange & Act
            var config1 = _serviceProvider.GetService<IConfigurationService>();
            var config2 = _serviceProvider.GetService<IConfigurationService>();

            // Assert
            Assert.Same(config1, config2);
        }

        [Fact]
        public void RegisterServices_TransientServices_ShouldReturnDifferentInstances()
        {
            // Arrange & Act
            var pdf1 = _serviceProvider.GetService<IPdfOperationsService>();
            var pdf2 = _serviceProvider.GetService<IPdfOperationsService>();

            // Assert
            Assert.NotSame(pdf1, pdf2);
        }

        [Fact]
        public void RegisterServices_ScopedServices_ShouldReturnSameInstanceInScope()
        {
            // Arrange
            using var scope = _serviceProvider.CreateScope();

            // Act
            var service1 = scope.ServiceProvider.GetService<IFileProcessingService>();
            var service2 = scope.ServiceProvider.GetService<IFileProcessingService>();

            // Assert
            Assert.Same(service1, service2);
        }

        [Fact]
        public void RegisterServices_ScopedServices_ShouldReturnDifferentInstancesAcrossScopes()
        {
            // Arrange & Act
            IFileProcessingService service1, service2;
            
            using (var scope1 = _serviceProvider.CreateScope())
            {
                service1 = scope1.ServiceProvider.GetService<IFileProcessingService>()!;
            }
            
            using (var scope2 = _serviceProvider.CreateScope())
            {
                service2 = scope2.ServiceProvider.GetService<IFileProcessingService>()!;
            }

            // Assert
            Assert.NotSame(service1, service2);
        }

        [Fact]
        public void RegisterServices_AllRequiredServices_ShouldResolveWithoutException()
        {
            // This test ensures all services can be constructed with their dependencies
            var exceptions = new List<Exception>();

            try { _serviceProvider.GetRequiredService<IConfigurationService>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            try { _serviceProvider.GetRequiredService<IProcessManager>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            try { _serviceProvider.GetRequiredService<IPerformanceMonitor>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            try { _serviceProvider.GetRequiredService<IPdfCacheService>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            try { _serviceProvider.GetRequiredService<IScopeOfWorkService>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            try { _serviceProvider.GetRequiredService<ICompanyNameService>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            try { _serviceProvider.GetRequiredService<IPdfOperationsService>(); }
            catch (Exception ex) { exceptions.Add(ex); }

            // Assert
            Assert.Empty(exceptions);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                _serviceProvider?.Dispose();
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}