# Phase 2 Completion Roadmap

**Project**: DocHandler Enterprise - Phase 2 Architecture Refactoring  
**Remaining Work**: ~30% of Phase 2 objectives  
**Estimated Duration**: 4-6 weeks  
**Priority**: Complete architectural foundation before Phase 3

---

## 🎯 Executive Summary

This roadmap outlines the remaining work items needed to complete Phase 2 of the DocHandler Enterprise modernization. All items are ordered by dependency and priority, with detailed implementation steps for each task.

---

## 📋 Week 1: Complete Dependency Injection Infrastructure

### Day 1-2: Extract Remaining Service Interfaces

**Objective**: Eliminate all direct instantiation in MainViewModel and throughout the application.

#### 1.1 Create IOfficeConversionService Interface
```csharp
// Location: Services/IOfficeConversionService.cs
public interface IOfficeConversionService : IDisposable
{
    bool IsOfficeInstalled();
    Task<bool> ConvertWordToPdfAsync(string inputPath, string outputPath);
    Task<bool> ConvertExcelToPdfAsync(string inputPath, string outputPath);
    ConversionResult ConvertWordToPdf(string inputPath, string outputPath);
    ConversionResult ConvertExcelToPdf(string inputPath, string outputPath);
}
```

#### 1.2 Update OfficeConversionService
- Add interface implementation
- Ensure all methods match interface signature
- Verify COM cleanup patterns remain intact

#### 1.3 Register in ServiceRegistration
```csharp
// Add to ServiceRegistration.cs
services.AddTransient<IOfficeConversionService, OfficeConversionService>();
```

#### 1.4 Update MainViewModel Constructor
- Add `IOfficeConversionService` parameter
- Remove `new OfficeConversionService()` instantiation
- Update null checks and initialization

### Day 3-4: Implement Mode-Specific Service Containers

**Objective**: Enable service isolation between modes as designed in Phase 2 investigation.

#### 1.5 Create IModeServiceProvider
```csharp
// Location: Services/IModeServiceProvider.cs
public interface IModeServiceProvider
{
    IServiceProvider GetModeServices(string modeName);
    T GetRequiredService<T>(string modeName) where T : class;
    object GetRequiredService(string modeName, Type serviceType);
    void RegisterModeServices(string modeName, Action<IServiceCollection> configure);
}
```

#### 1.6 Implement ModeServiceProvider
```csharp
// Location: Services/ModeServiceProvider.cs
public class ModeServiceProvider : IModeServiceProvider
{
    private readonly Dictionary<string, IServiceProvider> _modeProviders;
    private readonly IServiceProvider _rootProvider;
    
    // Implementation details...
}
```

#### 1.7 Update ModeRegistry
- Integrate with ModeServiceProvider
- Allow modes to register their specific services
- Maintain backward compatibility

### Day 5: Service Factory Pattern Implementation

#### 1.8 Create Service Factories
```csharp
// Location: Services/Factories/IOfficeServiceFactory.cs
public interface IOfficeServiceFactory
{
    T CreateService<T>() where T : IOfficeService;
    void ReleaseService<T>(T service) where T : IOfficeService;
}
```

#### 1.9 Update SaveQuotesMode
- Use service factory for Office services
- Ensure proper disposal through factory
- Test memory management remains intact

---

## 📋 Week 2: MVVM Refactoring - Part 1 (Business Logic Extraction)

### Day 1-2: Create Business Logic Services

**Objective**: Extract all business logic from MainViewModel into dedicated services.

#### 2.1 Create IFileProcessingOrchestrator
```csharp
// Location: Services/IFileProcessingOrchestrator.cs
public interface IFileProcessingOrchestrator
{
    Task<ProcessingResult> ProcessFilesAsync(
        IEnumerable<FileItem> files, 
        ProcessingOptions options,
        IProgress<ProcessingProgress> progress,
        CancellationToken cancellationToken);
        
    Task<ValidationResult> ValidateFilesAsync(IEnumerable<FileItem> files);
    Task<string> DetermineOutputDirectoryAsync(ProcessingOptions options);
}
```

#### 2.2 Extract File Processing Logic
- Move 500+ lines of processing logic from MainViewModel
- Maintain all existing functionality
- Ensure progress reporting continues to work
- Preserve error handling behavior

#### 2.3 Create IUserInteractionService
```csharp
// Location: Services/IUserInteractionService.cs
public interface IUserInteractionService
{
    Task<bool> ConfirmActionAsync(string message, string title);
    Task ShowErrorAsync(string message, string title, Exception? exception = null);
    Task ShowSuccessAsync(string message, string title);
    Task<string?> SelectFolderAsync(string? initialDirectory = null);
    Task<IEnumerable<string>> SelectFilesAsync(FileDialogOptions options);
}
```

### Day 3-4: Create Mode Coordination Service

#### 2.4 Create IModeCoordinationService
```csharp
// Location: Services/IModeCoordinationService.cs
public interface IModeCoordinationService
{
    string CurrentMode { get; }
    event EventHandler<ModeChangedEventArgs> ModeChanged;
    
    Task<bool> ActivateModeAsync(string modeName);
    Task<bool> DeactivateModeAsync();
    Task<IEnumerable<ModeAction>> GetModeActionsAsync();
    Task<ProcessingResult> ExecuteModeActionAsync(string actionName, object parameters);
}
```

#### 2.5 Implement ModeCoordinationService
- Coordinate between ModeManager and UI
- Handle mode switching logic
- Manage mode-specific UI updates
- Integrate with existing SaveQuotesMode

### Day 5: Update MainViewModel - Phase 1

#### 2.6 Refactor MainViewModel Constructor
- Inject new business services
- Remove business logic methods
- Update command implementations to use services
- Maintain all data binding properties

---

## 📋 Week 3: MVVM Refactoring - Part 2 (ViewModel Decomposition)

### Day 1-2: Create Specialized ViewModels

#### 3.1 Create FileListViewModel
```csharp
// Location: ViewModels/FileListViewModel.cs
public partial class FileListViewModel : ObservableObject
{
    [ObservableProperty]
    private ObservableCollection<FileItem> pendingFiles = new();
    
    [RelayCommand]
    private async Task AddFilesAsync();
    
    [RelayCommand]
    private void RemoveFile(FileItem file);
    
    [RelayCommand]
    private void ClearFiles();
}
```

#### 3.2 Create ProcessingViewModel
```csharp
// Location: ViewModels/ProcessingViewModel.cs
public partial class ProcessingViewModel : ObservableObject
{
    [ObservableProperty]
    private bool isProcessing;
    
    [ObservableProperty]
    private double progressValue;
    
    [ObservableProperty]
    private string statusMessage = "Ready";
    
    [RelayCommand]
    private async Task ProcessFilesAsync();
}
```

#### 3.3 Create ModeSelectionViewModel
- Handle mode switching UI
- Manage mode-specific options
- Coordinate with ModeCoordinationService

### Day 3-4: Implement Composite MainViewModel

#### 3.4 Refactor MainViewModel as Coordinator
```csharp
public partial class MainViewModel : ObservableObject
{
    // Child ViewModels
    public FileListViewModel FileList { get; }
    public ProcessingViewModel Processing { get; }
    public ModeSelectionViewModel ModeSelection { get; }
    public ConfigurationViewModel Configuration { get; }
    
    // Coordination logic only
}
```

#### 3.5 Update XAML Bindings
- Update all DataContext references
- Modify binding paths for child ViewModels
- Ensure commands still work correctly
- Test all UI functionality

### Day 5: Integration Testing

#### 3.6 Comprehensive UI Testing
- Verify all bindings work correctly
- Test mode switching functionality
- Ensure progress reporting works
- Validate error display mechanisms

---

## 📋 Week 4: Formal Testing Infrastructure

### Day 1-2: Create Test Project Structure

#### 4.1 Create Test Projects
```
Solution/
├── DocHandler.Tests.Unit/
│   ├── Services/
│   ├── ViewModels/
│   └── Helpers/
├── DocHandler.Tests.Integration/
│   ├── Modes/
│   ├── Pipeline/
│   └── Services/
└── DocHandler.Tests.Common/
    ├── Builders/
    ├── Fixtures/
    └── Mocks/
```

#### 4.2 Add Test Dependencies
```xml
<PackageReference Include="xunit" Version="2.6.1" />
<PackageReference Include="xunit.runner.visualstudio" Version="2.5.3" />
<PackageReference Include="Moq" Version="4.20.69" />
<PackageReference Include="FluentAssertions" Version="6.12.0" />
<PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.8.0" />
```

### Day 3-4: Implement Test Base Classes

#### 4.3 Create ModeTestBase
```csharp
// Location: DocHandler.Tests.Common/ModeTestBase.cs
public abstract class ModeTestBase<TMode> where TMode : class, IProcessingMode
{
    protected IServiceProvider Services { get; private set; }
    protected TMode Mode { get; private set; }
    protected Mock<ILogger> LoggerMock { get; private set; }
    
    [SetUp]
    public virtual void Setup()
    {
        var services = new ServiceCollection();
        ConfigureServices(services);
        Services = services.BuildServiceProvider();
        Mode = CreateMode();
    }
    
    protected abstract TMode CreateMode();
    protected virtual void ConfigureServices(IServiceCollection services) 
    {
        // Add common test services
    }
}
```

#### 4.4 Create Mock Implementations
```csharp
// Location: DocHandler.Tests.Common/Mocks/MockOfficeService.cs
public class MockOfficeService : IOfficeConversionService
{
    public bool SimulateFailure { get; set; }
    public int CallCount { get; private set; }
    
    public Task<bool> ConvertWordToPdfAsync(string inputPath, string outputPath)
    {
        CallCount++;
        if (SimulateFailure) return Task.FromResult(false);
        
        // Simulate successful conversion
        File.Copy(inputPath, outputPath);
        return Task.FromResult(true);
    }
}
```

### Day 5: Migrate Existing Tests

#### 4.5 Convert QuickDiagnostic Tests
- Extract memory leak tests to unit tests
- Convert thread safety tests
- Formalize mode infrastructure tests
- Add assertions using FluentAssertions

#### 4.6 Create Integration Tests
```csharp
// Location: DocHandler.Tests.Integration/Modes/SaveQuotesModeTests.cs
public class SaveQuotesModeTests : ModeTestBase<SaveQuotesMode>
{
    [Test]
    public async Task ProcessAsync_WithValidFiles_ProcessesSuccessfully()
    {
        // Arrange
        var files = new[] { CreateTestFile("test.docx") };
        var request = new ProcessingRequest { Files = files };
        
        // Act
        var result = await Mode.ProcessAsync(request, CancellationToken.None);
        
        // Assert
        result.Success.Should().BeTrue();
        result.ProcessedFiles.Should().HaveCount(1);
    }
}
```

---

## 📋 Week 5: Configuration System Completion

### Day 1-2: Implement Hot-Reload

#### 5.1 Create FileWatcher for Configuration
```csharp
// Location: Services/Configuration/ConfigurationFileWatcher.cs
public class ConfigurationFileWatcher : IDisposable
{
    private FileSystemWatcher _watcher;
    
    public event EventHandler<FileChangedEventArgs> ConfigurationChanged;
    
    public void StartWatching(string configPath)
    {
        _watcher = new FileSystemWatcher(Path.GetDirectoryName(configPath));
        _watcher.Filter = Path.GetFileName(configPath);
        _watcher.Changed += OnConfigurationFileChanged;
        _watcher.EnableRaisingEvents = true;
    }
}
```

#### 5.2 Integrate with HierarchicalConfigurationService
- Add file watcher integration
- Implement configuration reload logic
- Add debouncing for multiple changes
- Notify all subscribers of changes

### Day 3-4: Mode-Specific Configuration UI

#### 5.3 Create Configuration ViewModels
```csharp
// Location: ViewModels/Configuration/ModeConfigurationViewModel.cs
public abstract class ModeConfigurationViewModel<TConfig> : ObservableObject 
    where TConfig : class, new()
{
    protected TConfig Configuration { get; private set; }
    
    [RelayCommand]
    private async Task SaveConfigurationAsync();
    
    [RelayCommand]
    private void ResetToDefaults();
}
```

#### 5.4 Implement SaveQuotesConfigurationView
- Create WPF UserControl for configuration
- Bind to mode-specific settings
- Implement validation rules
- Add to Settings dialog

### Day 5: Configuration Migration Tool

#### 5.5 Create Configuration Migrator
```csharp
// Location: Services/Configuration/ConfigurationMigrator.cs
public class ConfigurationMigrator
{
    public async Task<MigrationResult> MigrateConfigurationAsync(
        string fromVersion, 
        string toVersion,
        HierarchicalAppConfiguration config)
    {
        // Implement version-specific migrations
    }
}
```

---

## 📋 Week 6: Final Integration and Polish

### Day 1-2: Service Integration Completion

#### 6.1 Update App.xaml.cs
- Ensure all services are registered
- Verify service lifetimes are correct
- Add logging for service creation
- Test service resolution

#### 6.2 Complete Mode Registration
```csharp
// Update ServiceRegistration.cs
public static IServiceCollection RegisterAllModes(this IServiceCollection services)
{
    services.RegisterProcessingMode<SaveQuotesMode>();
    // Future modes will be added here
    
    // Register mode-specific services
    services.RegisterModeServices<SaveQuotesMode>(modeServices =>
    {
        modeServices.AddTransient<ISaveQuotesValidator, SaveQuotesValidator>();
        // Other mode-specific services
    });
    
    return services;
}
```

### Day 3-4: Performance Validation

#### 6.3 Memory Leak Verification
- Run extended memory tests
- Verify COM cleanup still works
- Check for service disposal issues
- Monitor mode switching memory impact

#### 6.4 Threading Verification
- Ensure all async patterns are correct
- Verify UI responsiveness
- Check for deadlock scenarios
- Validate STA thread usage

### Day 5: Documentation and Handoff

#### 6.5 Update Technical Documentation
- Document all new services
- Create mode development guide
- Update architecture diagrams
- Write migration guide

#### 6.6 Create Developer Guide
```markdown
# Mode Development Guide

## Creating a New Mode
1. Inherit from ProcessingModeBase
2. Implement required methods
3. Register in ServiceRegistration
4. Create mode-specific UI (optional)
5. Add configuration section
6. Write unit tests
```

---

## 🎯 Success Criteria

### Technical Criteria
- [ ] Zero direct instantiation in application
- [ ] MainViewModel under 500 lines
- [ ] 80% unit test coverage
- [ ] All services use DI
- [ ] Hot-reload configuration working
- [ ] Mode isolation implemented

### Quality Criteria
- [ ] No memory leaks introduced
- [ ] No threading violations
- [ ] No performance degradation
- [ ] All existing features work
- [ ] Clean architecture maintained

### Documentation
- [ ] All services documented
- [ ] Mode development guide complete
- [ ] Test examples provided
- [ ] Architecture diagrams updated

---

## 🚀 Next Steps (Phase 3 Preparation)

Upon completion of Phase 2:
1. **Performance Profiling**: Baseline current performance
2. **Mode Development**: Plan additional modes
3. **UI Enhancement**: Design mode-specific UI components
4. **Plugin Architecture**: Design external mode loading
5. **Cloud Integration**: Plan remote processing capabilities

---

## 📅 Timeline Summary

- **Week 1**: Complete DI Infrastructure (Critical Path)
- **Week 2**: MVVM Refactoring Part 1 (High Priority)
- **Week 3**: MVVM Refactoring Part 2 (High Priority)
- **Week 4**: Testing Infrastructure (Medium Priority)
- **Week 5**: Configuration Completion (Medium Priority)
- **Week 6**: Integration and Polish (Final Validation)

**Total Duration**: 6 weeks  
**Risk Buffer**: +1 week for unforeseen issues  
**Recommended Team Size**: 2-3 developers  

---

*This roadmap focuses exclusively on completing the remaining 30% of Phase 2 work. All items are ordered by dependency and priority to minimize risk and ensure architectural stability.*