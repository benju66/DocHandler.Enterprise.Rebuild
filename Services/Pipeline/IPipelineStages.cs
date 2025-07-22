using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services.Pipeline
{
    /// <summary>
    /// Base interface for all pipeline stages
    /// </summary>
    public interface IPipelineStage
    {
        string StageName { get; }
        Task<bool> CanProcessAsync(FileItem file, ProcessingContext context);
    }

    /// <summary>
    /// File validation stage - validates files before processing
    /// </summary>
    public interface IFileValidator : IPipelineStage
    {
        Task<ValidationResult> ValidateAsync(FileItem file, ProcessingContext context);
    }

    /// <summary>
    /// Pre-processing stage - extracts information, prepares files
    /// </summary>
    public interface IPreProcessor : IPipelineStage
    {
        Task<PreProcessingResult> ProcessAsync(FileItem file, ProcessingContext context);
    }

    /// <summary>
    /// Conversion stage - main file processing/conversion
    /// </summary>
    public interface IFileConverter : IPipelineStage
    {
        Task<ConversionResult> ConvertAsync(FileItem file, ProcessingContext context);
    }

    /// <summary>
    /// Post-processing stage - optimizes, organizes, finalizes output
    /// </summary>
    public interface IPostProcessor : IPipelineStage
    {
        Task<PostProcessingResult> ProcessAsync(ConversionResult input, ProcessingContext context);
    }

    /// <summary>
    /// Output generation stage - creates final organized output
    /// </summary>
    public interface IOutputGenerator : IPipelineStage
    {
        Task<OutputResult> GenerateAsync(List<PostProcessingResult> inputs, ProcessingContext context);
    }

    /// <summary>
    /// Result from pre-processing stage
    /// </summary>
    public class PreProcessingResult
    {
        public bool Success { get; set; }
        public FileItem ProcessedFile { get; set; }
        public Dictionary<string, object> ExtractedData { get; set; } = new();
        public List<string> Messages { get; set; } = new();
        public Exception Error { get; set; }
    }

    /// <summary>
    /// Result from conversion stage
    /// </summary>
    public class ConversionResult
    {
        public bool Success { get; set; }
        public string OutputPath { get; set; }
        public FileItem SourceFile { get; set; }
        public Dictionary<string, object> ConversionData { get; set; } = new();
        public List<string> Messages { get; set; } = new();
        public Exception Error { get; set; }
        public TimeSpan ProcessingTime { get; set; }
    }

    /// <summary>
    /// Result from post-processing stage
    /// </summary>
    public class PostProcessingResult
    {
        public bool Success { get; set; }
        public string FinalPath { get; set; }
        public ConversionResult SourceConversion { get; set; }
        public Dictionary<string, object> PostProcessingData { get; set; } = new();
        public List<string> Messages { get; set; } = new();
        public Exception Error { get; set; }
    }

    /// <summary>
    /// Result from output generation stage
    /// </summary>
    public class OutputResult
    {
        public bool Success { get; set; }
        public List<string> OutputPaths { get; set; } = new();
        public Dictionary<string, object> OutputMetadata { get; set; } = new();
        public List<string> Messages { get; set; } = new();
        public Exception Error { get; set; }
    }
} 