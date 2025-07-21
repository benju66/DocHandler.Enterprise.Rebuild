# Phase 1: Critical Stabilization with Mode Foundation Investigation Guide
**Duration**: Weeks 2-5
**Objective**: Fix critical issues while building mode infrastructure foundation

## Week 2: COM Memory Management + Abstraction Layer

### Investigation Tasks

#### 1. COM Memory Leak Analysis
**Files to examine**:
- `Services/OfficeConversionService.cs`
- `Services/SessionAwareOfficeService.cs`
- `Services/SessionAwareExcelService.cs`
- `Services/ReliableOfficeConverter.cs`

**Look for**:
```csharp
// MEMORY LEAK PATTERN - Documents/Workbooks collections not released
doc = _wordApp.Documents.Open(...); // ❌ Collection reference not tracked
workbook = _excelApp.Workbooks.Open(...); // ❌ Collection reference not tracked

// Should be:
using (var comScope = new ComResourceScope())
{
    var documents = comScope.Track(_wordApp.Documents, "Documents", "Context");
    var doc = comScope.Track(documents.Open(...), "Document", "Context");
}
```

**Action Items**:
1. Find all COM collection access points
2. Verify Marshal.ReleaseComObject usage
3. Check for missing finally blocks
4. Identify missing using statements
5. Document current memory growth patterns

#### 2. Create Mode-Agnostic COM Layer
**Design interfaces**:
```csharp
public interface IDocumentConverter
{
    bool CanConvert(string fromFormat, string toFormat);
    Task<ConversionResult> ConvertAsync(ConversionRequest request);
}

public interface IOfficeServiceFactory
{
    T CreateService<T>() where T : IOfficeService;
    void ReleaseService<T>(T service) where T : IOfficeService;
}
```

**Implementation requirements**:
- Thread-safe service creation
- Automatic COM cleanup
- Mode-specific options support
- Performance metrics collection

## Week 3: Threading Model + Mode Execution Context

### Investigation Tasks

#### 1. Threading Violation Analysis
**Files to examine**:
- `Services/SaveQuotesQueueService.cs` - Critical STA violations
- `Services/StaThreadPool.cs` - Custom implementation issues
- `Services/OptimizedFileProcessingService.cs` - Task.Run usage

**Look for**:
```csharp
// ❌ WRONG - Creates MTA thread
await Task.Run(() => 
{
    _wordApp.Documents.Open(...); // COM call fails
});

// ✅ CORRECT - Uses STA thread
await _staThreadPool.ExecuteAsync(() =>
{
    _wordApp.Documents.Open(...); // COM call succeeds
});
```

**Critical Fixes**:
1. Replace ALL Task.Run with STA-aware execution
2. Add ConfigureAwait(false) to prevent deadlocks
3. Fix race conditions in queue processing
4. Implement proper cancellation token propagation

#### 2. Mode Execution Framework
**Create foundation for**:
```csharp
public class ModeExecutionContext
{
    public string ModeName { get; }
    public IServiceProvider Services { get; }
    public CancellationToken CancellationToken { get; }
    public ILogger Logger { get; }
    public IMetricsCollector Metrics { get; }
}
```

## Week 4: Security Hardening + Mode Permissions

### Investigation Tasks

#### 1. Security Vulnerability Audit
**Files to examine**:
- `Helpers/FileValidator.cs` - Path traversal checks
- `Services/OptimizedFileProcessingService.cs` - File operations
- `ViewModels/MainViewModel.cs` - User input handling

**Security checklist**:
- [ ] Path canonicalization implemented
- [ ] File content validation (not just extensions)
- [ ] Input sanitization for all user data
- [ ] Macro execution disabled
- [ ] File size limits enforced
- [ ] Temporary file cleanup verified

#### 2. Mode Security Framework
**Implement**:
```csharp
public interface IModeSecurityPolicy
{
    bool ValidateFilePath(string path);
    bool ValidateFileContent(byte[] content, string extension);
    long GetMaxFileSize(string fileType);
    IEnumerable<string> GetAllowedExtensions();
    bool RequiresElevation();
}
```

## Week 5: Error Handling Foundation

### Investigation Tasks

#### 1. Exception Management Audit
**Look for**:
```csharp
// ❌ WRONG - Swallowed exception
catch (Exception ex)
{
    _logger.Warning(ex, "Error");
    // Continues with corrupted state
}

// ✅ CORRECT - Proper handling
catch (COMException ex) when (ex.HResult == 0x800A03EC)
{
    _logger.Error(ex, "Excel automation failed");
    throw new ModeOperationException("Excel conversion failed", ex)
    {
        ErrorCode = "EXCEL_AUTOMATION_FAIL",
        RecoveryAction = RecoveryAction.RetryWithNewInstance
    };
}
```

#### 2. Correlation ID Implementation
**Add to all operations**:
```csharp
public class OperationContext
{
    public string CorrelationId { get; } = Guid.NewGuid().ToString();
    public string ModeName { get; set; }
    public Dictionary<string, object> Properties { get; } = new();
}
```

## Testing Requirements

### Memory Leak Tests
```csharp
[Test]
public async Task ConvertDocument_DisposesComObjects()
{
    // Arrange
    var initialCount = ComHelper.GetComObjectCount();
    
    // Act
    await ConvertDocumentAsync();
    GC.Collect();
    GC.WaitForPendingFinalizers();
    GC.Collect();
    
    // Assert
    var finalCount = ComHelper.GetComObjectCount();
    Assert.AreEqual(initialCount, finalCount);
}
```

### Threading Tests
```csharp
[Test]
public async Task QueueProcessing_UsesStaThreads()
{
    // Arrange
    var queue = new SaveQuotesQueueService(...);
    
    // Act
    await queue.ProcessAsync(() =>
    {
        // Assert
        Assert.AreEqual(ApartmentState.STA, 
            Thread.CurrentThread.GetApartmentState());
    });
}
```

## Deliverables

1. **Fixed Code Files**
   - All COM leaks patched
   - Threading violations resolved
   - Security vulnerabilities fixed
   - Error handling implemented

2. **New Abstraction Layer**
   - IDocumentConverter interface
   - IModeExecutor implementation
   - IModeSecurityPolicy framework
   - Error handling infrastructure

3. **Test Suite**
   - Memory leak detection tests
   - Threading verification tests
   - Security validation tests
   - Error recovery tests

4. **Documentation**
   - COM best practices guide
   - Threading model documentation
   - Security implementation guide
   - Error handling patterns

## Critical Warnings

⚠️ **COM Operations**:
- NEVER access Office objects from non-STA threads
- ALWAYS release collections before documents
- MUST use try-finally for cleanup

⚠️ **Threading**:
- NO Task.Run for COM operations
- ALWAYS use ConfigureAwait(false)
- MUST handle cancellation properly

⚠️ **Security**:
- NEVER trust file extensions alone
- ALWAYS canonicalize paths
- MUST validate file content

## Success Criteria
- [ ] Zero COM memory leaks in 24-hour test
- [ ] All operations run on correct threads
- [ ] Security scan shows no vulnerabilities
- [ ] 100% of exceptions properly handled
- [ ] Mode abstraction layer functional
- [ ] All tests passing