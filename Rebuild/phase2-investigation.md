# Phase 2: Architecture Refactoring for Modes Investigation Guide
**Duration**: Weeks 6-10
**Objective**: Build sustainable architecture supporting multiple modes

## Week 6: Dependency Injection + Mode Registry

### Investigation Tasks

#### 1. Current DI Analysis
**Files to examine**:
- `Program.cs` - Service registration
- `Services/ServiceRegistration.cs` - Existing DI setup
- `ViewModels/MainViewModel.cs` - Direct instantiation anti-patterns

**Identify problems**:
```csharp
// ❌ WRONG - Direct instantiation in MainViewModel
_sessionOfficeService = new SessionAwareOfficeService();
_fileProcessingService = new OptimizedFileProcessingService(...);

// ✅ CORRECT - Constructor injection
public MainViewModel(
    ISessionOfficeService sessionOfficeService,
    IFileProcessingService fileProcessingService)
{
    _sessionOfficeService = sessionOfficeService;
    _fileProcessingService = fileProcessingService;
}
```

#### 2. Implement Mode Registry
**Create infrastructure**:
```csharp
public interface IModeRegistry
{
    void Register<TMode>(ServiceLifetime lifetime = ServiceLifetime.Scoped) 
        where TMode : class, IProcessingMode;
    
    void RegisterModeServices(string modeName, Action<IServiceCollection> configure);
    
    IProcessingMode CreateMode(string modeName, IServiceProvider serviceProvider);
    
    IEnumerable<ModeDescriptor> GetAvailableModes();
}

public class ModeDescriptor
{
    public string Name { get; set; }
    public string DisplayName { get; set; }
    public string Description { get; set; }
    public Type ModeType { get; set; }
    public ModeCapabilities Capabilities { get; set; }
}
```

#### 3. Service Isolation Strategy
**Design mode-specific containers**:
```csharp
public interface IModeServiceProvider
{
    IServiceProvider GetModeServices(string modeName);
    T GetRequiredService<T>(string modeName);
    object GetRequiredService(string modeName, Type serviceType);
}
```

### Implementation Steps
1. Extract all service interfaces
2. Create mode-aware service factory
3. Implement service lifetime management
4. Add service validation
5. Create service composition root

## Week 7: MVVM Refactoring + Mode UI Framework

### Investigation Tasks

#### 1. ViewModel Analysis
**Files to examine**:
- `ViewModels/MainViewModel.cs` - 1000+ lines of mixed concerns
- `ViewModels/SettingsViewModel.cs` - Business logic in VM
- All ViewModels for business logic violations

**Extract to services**:
```csharp
// ❌ WRONG - Business logic in ViewModel
public partial class MainViewModel
{
    private async Task ConvertFiles()
    {
        // 100+ lines of conversion logic
    }
}

// ✅ CORRECT - Logic in service
public partial class MainViewModel
{
    private readonly IFileConversionService _conversionService;
    
    [RelayCommand]
    private async Task ConvertFiles()
    {
        await _conversionService.ConvertAsync(Files);
    }
}
```

#### 2. Mode UI Framework Design
**Create dynamic UI system**:
```csharp
public interface IModeUIProvider
{
    // Core UI components
    UserControl GetMainPanel();
    IEnumerable<MenuItem> GetMenuItems();
    IEnumerable<ToolBarItem> GetToolBarItems();
    
    // Dynamic UI updates
    void UpdateUIState(ModeState state);
    void RegisterUIExtension(IModeUIExtension extension);
    
    // Validation
    IEnumerable<ValidationRule> GetValidationRules();
}

public interface IModeViewModel : INotifyPropertyChanged
{
    string ModeName { get; }
    ICommand ProcessCommand { get; }
    bool CanProcess { get; }
    void Initialize(IModeContext context);
}
```

#### 3. UI Composition Strategy
**Implement mode switching**:
```csharp
public class ModeUIManager
{
    public void SwitchToMode(string modeName)
    {
        // 1. Cleanup current mode UI
        // 2. Load new mode UI
        // 3. Update ribbon/menu
        // 4. Apply mode theme
        // 5. Restore mode state
    }
}
```

## Week 8: Mode Processing Pipeline

### Investigation Tasks

#### 1. Current Processing Analysis
**Map the flow**:
- File selection → Validation → Processing → Output
- Identify common patterns
- Find mode-specific branches
- Document decision points

#### 2. Pipeline Architecture Implementation
**Create extensible pipeline**:
```csharp
public interface IPipelineBuilder
{
    IPipelineBuilder UseValidator<TValidator>() 
        where TValidator : IFileValidator;
    
    IPipelineBuilder UsePreProcessor<TProcessor>() 
        where TProcessor : IPreProcessor;
    
    IPipelineBuilder UseConverter<TConverter>() 
        where TConverter : IFileConverter;
    
    IPipelineBuilder UsePostProcessor<TProcessor>() 
        where TProcessor : IPostProcessor;
    
    IProcessingPipeline Build();
}

public class ProcessingContext
{
    public string CorrelationId { get; }
    public IReadOnlyList<FileItem> InputFiles { get; }
    public Dictionary<string, object> Properties { get; }
    public CancellationToken CancellationToken { get; }
    public IProgress<ProcessingProgress> Progress { get; }
}
```

#### 3. Migrate Save Quotes Mode
**Steps**:
1. Create `SaveQuotesMode : IProcessingMode`
2. Extract validation logic to `SaveQuotesValidator`
3. Extract conversion logic to `SaveQuotesConverter`
4. Create mode-specific pipeline configuration
5. Maintain backward compatibility

## Week 9: Testing Infrastructure

### Investigation Tasks

#### 1. Test Framework Setup
**Create test hierarchy**:
```csharp
public abstract class ModeTestBase<TMode> where TMode : IProcessingMode
{
    protected IServiceProvider Services { get; private set; }
    protected TMode Mode { get; private set; }
    protected IModeContext Context { get; private set; }
    
    [SetUp]
    public virtual void Setup()
    {
        var services = new ServiceCollection();
        ConfigureServices(services);
        Services = services.BuildServiceProvider();
        Mode = CreateMode();
        Context = CreateContext();
    }
    
    protected abstract TMode CreateMode();
    protected abstract IModeContext CreateContext();
    protected virtual void ConfigureServices(IServiceCollection services) { }
}
```

#### 2. Test Categories
**Implement tests for**:
- Unit tests for each service
- Integration tests for mode workflows
- Performance benchmarks
- Memory leak detection
- Thread safety verification
- Security validation

#### 3. Mock Infrastructure
**Create test doubles**:
```csharp
public class MockOfficeService : IOfficeService
{
    public Task<ConversionResult> ConvertAsync(ConversionRequest request)
    {
        // Simulate conversion without Office
    }
}

public class TestFileSystem : IFileSystem
{
    private readonly Dictionary<string, byte[]> _files = new();
    
    public void AddFile(string path, byte[] content)
    {
        _files[path] = content;
    }
}
```

## Week 10: Configuration System

### Investigation Tasks

#### 1. Current Configuration Analysis
**Files to examine**:
- `Services/ConfigurationService.cs`
- `AppConfiguration` class
- All hardcoded values

#### 2. Hierarchical Configuration Design
**Implement structure**:
```yaml
# Global settings
Application:
  Theme: Dark
  LogLevel: Information
  Culture: en-US

# Mode defaults
ModeDefaults:
  MaxFileSize: 50MB
  Timeout: 300s
  MaxConcurrency: 5

# Mode-specific overrides
Modes:
  SaveQuotes:
    MaxFileSize: 100MB
    EnableCompanyScanning: true
    CustomSettings:
      DefaultScope: "03-1000"
      
  BulkConversion:
    MaxConcurrency: 10
    OutputFormat: PDF
```

#### 3. Configuration Management
**Create infrastructure**:
```csharp
public interface IModeConfigurationManager
{
    T GetConfiguration<T>(string modeName) where T : class, new();
    void UpdateConfiguration<T>(string modeName, Action<T> update);
    IDisposable RegisterChangeCallback(string modeName, Action<object> callback);
}
```

## Deliverables

1. **Refactored Architecture**
   - Clean DI implementation
   - Mode registry system
   - Thin ViewModels
   - Pipeline architecture
   - Test infrastructure

2. **Mode Implementation**
   - SaveQuotesMode fully migrated
   - Mode UI framework functional
   - Configuration system working
   - Tests passing

3. **Documentation**
   - Architecture diagrams
   - Mode development guide
   - Testing best practices
   - Configuration schema

## Success Criteria
- [ ] Zero direct instantiation in ViewModels
- [ ] All services have interfaces
- [ ] Mode registry supports dynamic loading
- [ ] UI updates dynamically with mode
- [ ] Pipeline processes files correctly
- [ ] 70% test coverage achieved
- [ ] Configuration hot-reload works
- [ ] Save Quotes Mode fully migrated