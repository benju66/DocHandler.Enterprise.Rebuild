# Memory Leak Detector for DocHandler

## Overview

The Memory Leak Detector is a static analysis tool that identifies potential memory leaks and disposal issues in C# code, specifically targeting COM object management and Office automation patterns common in DocHandler.

## Features

### Detection Rules

The detector implements the following rules based on your configuration:

1. **Object Creation Without Using Statement** (Warning)
   - Pattern: `new\s+\w+\(\)(?!.*using\s*\()`
   - Detects object instantiation that may need proper disposal

2. **COM Collection Access** (Critical)
   - Pattern: `_\w+App\.\w+\.Open\([^)]*\)`
   - Identifies COM object access that requires disposal verification

3. **Office Collection Access** (Critical)
   - Pattern: `\.Documents\.|\.Workbooks\.`
   - Flags Office automation collection usage that needs tracking

4. **Event Subscription Without Unsubscribe** (Warning)
   - Pattern: `\+=(?!.*\-=)`
   - Finds event handlers that may cause memory leaks

### Smart Pattern Recognition

The detector includes intelligent filtering to reduce false positives:

- **Safe Patterns**: Recognizes `ComResourceScope` usage, `using` statements, and proper disposal patterns
- **Context Awareness**: Analyzes surrounding code for cleanup mechanisms
- **Value Type Filtering**: Excludes safe value types and framework collections

## Usage

### 1. From Within the Application

The memory leak detector is integrated into the MainViewModel with a `RunMemoryLeakAnalysisCommand`:

```csharp
// In the UI, bind to this command
await RunMemoryLeakAnalysisCommand.ExecuteAsync(null);
```

The analysis runs in the background and displays results in a popup window.

### 2. Standalone Console Application

Use the `MemoryLeakAnalyzer` class for command-line analysis:

```bash
# Analyze current directory
dotnet run MemoryLeakAnalyzer.cs

# Analyze specific directory
dotnet run MemoryLeakAnalyzer.cs "C:\Path\To\Project"

# Analyze specific file
dotnet run MemoryLeakAnalyzer.cs "Services\ReliableOfficeConverter.cs"
```

### 3. Programmatic Usage

```csharp
var detector = new MemoryLeakDetector();

// Analyze a single file
var fileResults = detector.AnalyzeFile("Services/ReliableOfficeConverter.cs");

// Analyze entire directory
var directoryResults = detector.AnalyzeDirectory("Services");

// Generate report
var report = detector.GenerateReport(directoryResults);
Console.WriteLine(report);
```

## Configuration

The detector uses configuration from `Data/memory-leak-detector-config.json`:

```json
{
  "agents": [
    {
      "name": "memory-leak-detector",
      "description": "Detects potential memory leaks and disposal issues",
      "enabled": true,
      "mode": "detect-only",
      "rules": [
        {
          "pattern": "new\\s+\\w+\\(\\)(?!.*using\\s*\\()",
          "message": "Object creation without using statement",
          "severity": "warning"
        }
      ]
    }
  ]
}
```

## Understanding Results

### Report Format

```
🔍 Memory Leak Detection Report
========================================

📊 Summary:
   Critical Issues: 2
   Warnings: 5
   Total Issues: 7

🚨 CRITICAL ISSUES:
------------------------------

📁 ReliableOfficeConverter.cs:
   Line 45: COM collection access - verify disposal
   Code: doc = _wordApp.Documents.Open(inputPath)

⚠️  WARNINGS:
--------------------

📁 MainViewModel.cs:
   Line 123: Object creation without using statement
   Code: var service = new SomeDisposableService()
```

### Severity Levels

- **Critical**: Issues that will likely cause memory leaks (COM objects, Office collections)
- **Warning**: Issues that may cause problems under certain conditions

## Integration with Existing Systems

### ComResourceScope Integration

The detector recognizes your existing `ComResourceScope` pattern as safe:

```csharp
// This will NOT trigger a warning
using (var comScope = new ComResourceScope())
{
    var documents = comScope.Track(_wordApp.Documents, "Documents");
    // Safe COM object usage
}
```

### Health Monitoring Integration

The detector works alongside your existing `OfficeHealthMonitor` and `OfficeProcessGuard` systems to provide comprehensive memory leak protection.

## Best Practices Detected

The analyzer promotes these patterns used in your codebase:

1. **Use ComResourceScope** for COM object management
2. **Wrap disposables in using statements**
3. **Unsubscribe from events** in disposal methods
4. **Implement IDisposable** for classes with unmanaged resources

## Continuous Integration

You can integrate the detector into your build process:

```xml
<!-- In your .csproj file -->
<Target Name="MemoryLeakAnalysis" BeforeTargets="Build">
  <Exec Command="dotnet run MemoryLeakAnalyzer.cs $(ProjectDir)" />
</Target>
```

## Troubleshooting

### False Positives

If the detector reports false positives:

1. **Check patterns**: Ensure your code follows established patterns like `ComResourceScope`
2. **Add comments**: The detector skips commented lines
3. **Review context**: The detector analyzes surrounding code context

### Performance

- **File size**: Large files may take longer to analyze
- **Regex complexity**: Complex patterns are pre-compiled for performance
- **Memory usage**: Analysis is done in-memory for speed

## Future Enhancements

Planned improvements:

1. **Custom Rules**: Load additional rules from configuration
2. **Fix Suggestions**: Automatic code fix proposals
3. **IDE Integration**: Visual Studio extension support
4. **Baseline Support**: Compare against previous analysis results

## Technical Details

### Architecture

```
MemoryLeakDetector
├── DetectionRule (regex patterns)
├── DetectionResult (findings)
├── IsSafePattern() (false positive filtering)
└── GenerateReport() (output formatting)
```

### Dependencies

- **Serilog**: Logging framework
- **System.Text.RegularExpressions**: Pattern matching
- **CommunityToolkit.Mvvm**: Command binding (for UI integration)

## Support

For issues or questions about the memory leak detector:

1. Check the analysis logs in `memory-leak-analysis.log`
2. Review the generated reports for context
3. Examine the `MEMORY_LEAK_FIX_SUMMARY.md` for background on memory management approaches