# COM Memory Leak Fixes Documentation

## Overview
This document describes the COM memory leak fixes implemented using the new `ComResourceScope` class.

## Problems Fixed

### 1. Documents/Workbooks Collection Leaks
Previously, accessing `wordApp.Documents` or `excelApp.Workbooks` created COM objects that were never released:
```csharp
// OLD - LEAKS Memory
doc = wordApp.Documents.Open(filePath);  // Documents collection leaked!
```

### 2. Health Check Memory Leaks
The periodic health check was creating and never releasing Documents collections:
```csharp
// OLD - LEAKS Memory every 5 minutes
var docCount = instance.Application.Documents.Count;  // Documents collection leaked!
```

### 3. Conditional Release Pattern Leaks
COM objects were only released if count > 0, leaking empty collections:
```csharp
// OLD - LEAKS when count = 0
var documents = _wordApp.Documents;
if (documents.Count > 0) {
    // ... cleanup ...
    ComHelper.SafeReleaseComObject(documents);  // Not called if count = 0!
}
```

## Solution: ComResourceScope

### Design Principles
- **Single Responsibility**: COM resource management only
- **RAII Pattern**: Automatic cleanup via IDisposable
- **Type Safety**: Tracks all COM objects with metadata
- **Minimal Refactoring**: Drop-in replacement for existing patterns

### Usage Examples

#### Opening Documents
```csharp
// NEW - Automatic cleanup
using (var comScope = new ComResourceScope())
{
    var doc = comScope.OpenWordDocument(wordApp, filePath);
    // Work with document...
} // All COM objects released here
```

#### Health Checks
```csharp
// NEW - No more leaks
using (var comScope = new ComResourceScope())
{
    var documents = comScope.GetDocuments(wordApp, "HealthCheck");
    var count = documents.Count;
} // Documents collection released
```

#### Disposal Patterns
```csharp
// NEW - Always releases collections
using (var comScope = new ComResourceScope())
{
    var documents = comScope.GetDocuments(wordApp, "Dispose");
    if (documents != null && documents.Count > 0)
    {
        // Close documents...
    }
} // Documents released regardless of count
```

## Files Updated

1. **Services/ComResourceScope.cs** (New)
   - Core COM resource management class

2. **Services/SessionAwareOfficeService.cs**
   - Fixed `OpenDocumentSafely()` - no more Documents leak
   - Fixed `DisposeWordApp()` - proper cleanup even for empty collections

3. **Services/SessionAwareExcelService.cs**
   - Fixed `ConvertSpreadsheetToPdf()` - no more Workbooks leak
   - Fixed `DisposeExcel()` - proper cleanup even for empty collections

4. **Services/OptimizedOfficeConversionService.cs**
   - Fixed `IsWordInstanceHealthy()` - no more periodic leaks
   - Fixed `OpenDocumentSafely()` - no more Documents leak
   - Fixed `DisposeWordInstance()` - proper cleanup

## Verification

Monitor COM object statistics using:
```csharp
ComHelper.LogComObjectStats();
```

Expected results after fixes:
- Balanced created/released counts
- No growing WINWORD.EXE processes
- Stable memory usage over time

## Best Practices Going Forward

1. **Always use ComResourceScope** for COM operations:
   ```csharp
   using (var comScope = new ComResourceScope())
   {
       // All COM operations here
   }
   ```

2. **Track all COM objects**:
   ```csharp
   var someObject = comScope.Track(app.SomeProperty, "PropertyType", "Context");
   ```

3. **Use helper methods** for common operations:
   - `OpenWordDocument()`
   - `OpenExcelWorkbook()`
   - `GetDocuments()`
   - `GetWorkbooks()`

## Migration Guide

To fix a COM leak in existing code:

1. Identify COM property access (e.g., `app.Documents`)
2. Wrap in `using (var comScope = new ComResourceScope())`
3. Use `comScope.Track()` or helper methods
4. Remove manual `ComHelper.SafeReleaseComObject()` calls within the scope

## Performance Impact

- **Minimal overhead**: Simple list tracking
- **Deterministic cleanup**: Happens at scope exit
- **No GC pressure**: Uses dispose pattern efficiently

## Future Improvements

- Add PowerPoint support when needed
- Add Outlook support if required
- Consider adding metrics/telemetry
- Potential for unit testing COM lifecycle

## Excel-Specific Fixes (Latest Update)

### 1. **Fixed Duplicate COM Tracking**
**Problem**: Workbook was being tracked twice - once by ComResourceScope and once manually
```csharp
// OLD - Double tracking caused stats mismatch
workbook = comScope.OpenExcelWorkbook(_excelApp, inputPath, readOnly: true);
ComHelper.TrackComObjectCreation("Workbook", "SessionAwareConvertToPdf"); // DUPLICATE!
```

**Fix**: Removed duplicate tracking since ComResourceScope already handles it

### 2. **Fixed WarmUp Race Condition**
**Problem**: Excel created in background thread might not be disposed if app closes quickly
```csharp
// OLD - No disposal tracking on warm-up failure
_sessionExcelService.WarmUp();
```

**Fix**: Added proper error handling and disposal on warm-up failure:
- Added warm-up lock to prevent multiple initializations
- Added disposal on exception
- Added COM stats logging after warm-up

### 3. **Enhanced Lifecycle Logging**
Added detailed logging to track Excel app creation and disposal:
- Thread ID logging to identify cross-thread issues
- Timestamp logging for lifecycle tracking
- COM stats logging after key operations

### 4. **Thread Safety Improvements**
- Added `_isWarmingUp` flag and `_warmUpLock` to prevent concurrent warm-ups
- Ensures only one Excel instance per service

## Monitoring Excel Leaks

Use the new COM Statistics menu item (Help → COM Object Statistics...) to:
1. Check for unreleased Excel instances
2. Monitor COM object balance
3. Track memory growth patterns

Expected healthy state:
- Net Objects: 0 (all created objects have been released)
- No "ExcelApp" entries in object details when idle
- Stable memory usage after processing files 

## Latest Fixes - Cross-Thread COM Access (December 2024)

### 1. **Fixed Cross-Thread COM Object Creation**
**Problem**: WarmUp methods were creating COM objects on MTA threads via Task.Run
```csharp
// OLD - Creates COM objects on MTA thread pool
_ = Task.Run(() =>
{
    _sessionOfficeService.WarmUp();
    _sessionExcelService.WarmUp();
});
```

**Fix**: Ensure COM objects are created on UI thread (STA)
```csharp
// FIXED - Creates COM objects on UI thread
if (Application.Current.Dispatcher.CheckAccess())
{
    _sessionOfficeService.WarmUp();
    _sessionExcelService.WarmUp();
}
else
{
    Application.Current.Dispatcher.BeginInvoke(new Action(() =>
    {
        PreWarmOfficeServicesForSaveQuotes();
    }));
}
```

### 2. **Fixed Health Check COM Object Creation**
**Problem**: Health checks were accessing COM properties that create hidden objects
```csharp
// OLD - Creates COM objects during health check
var _ = _excelApp.Version;
var documents = wordApp.Documents;
var docCount = documents.Count;
```

**Fix**: Simplified health checks to avoid property access
```csharp
// FIXED - No COM property access
if (_excelApp == null) return false;
return true; // Just check reference
```

### 3. **Fixed Service Disposal Order**
**Problem**: Services depending on Office were disposed before Office services
**Fix**: Dispose in correct order:
1. Queue service (uses file processing)
2. File processing service (uses Office)
3. Company name service (uses Office)
4. Office services last
5. Force COM cleanup
6. Log final stats

### 4. **Fixed WarmUp Property Access**
**Problem**: Setting many Word properties during warm-up creates COM objects
**Fix**: Only set essential properties:
```csharp
wordApp.Visible = false;
wordApp.DisplayAlerts = 0;
// Skip other properties
```

## Latest Fixes - Additional COM Memory Leaks (December 2024)

### 5. **Fixed foreach Loops Over COM Collections**
**Problem**: Using foreach on COM collections creates hidden enumerator objects
```csharp
// OLD - Creates COM enumerator that leaks
foreach (dynamic doc in documents)
{
    doc.Close(SaveChanges: false);
}
```

**Fix**: Use indexed for loops instead
```csharp
// FIXED - No enumerator created
int count = documents.Count;
for (int i = count; i >= 1; i--) // COM collections are 1-based
{
    dynamic doc = documents[i];
    doc.Close(SaveChanges: false);
    ComHelper.SafeReleaseComObject(doc, "Document", "Dispose");
}
```

**Fixed in:**
- OfficeConversionService.cs (Word and Excel disposal)
- SessionAwareOfficeService.cs (Word disposal)
- SessionAwareExcelService.cs (Excel disposal)
- OptimizedOfficeConversionService.cs (Word instance disposal)

### 6. **Fixed ConvertWordToPdfSync Orphaned Apps**
**Problem**: Nested try-catch blocks with early returns bypassed cleanup
```csharp
// OLD - Early returns bypass finally block
try
{
    wordApp = Activator.CreateInstance(wordType);
}
catch (Exception ex)
{
    return new ConversionResult { ... }; // Bypasses finally!
}
```

**Fix**: Removed nested try-catch, let exceptions bubble up
```csharp
// FIXED - Exceptions bubble to main try-catch
wordApp = Activator.CreateInstance(wordType);
if (wordApp == null)
{
    throw new InvalidOperationException("Failed to create Word application");
}
// Main finally block will always run
```

## Verification Steps
1. Run application and enable Save Quotes Mode
2. Check COM Statistics (Help → COM Object Statistics...)
3. Process several files
4. Check stats again - Net Objects should be 0
5. Close application - should close cleanly without hanging
6. Check Task Manager - no orphaned WINWORD.EXE or EXCEL.EXE processes 