using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using DocHandler.Models;

namespace DocHandler.Services.Pipeline
{
    /// <summary>
    /// Builder interface for creating processing pipelines
    /// </summary>
    public interface IPipelineBuilder
    {
        IPipelineBuilder UseValidator<TValidator>() where TValidator : IFileValidator;
        IPipelineBuilder UsePreProcessor<TProcessor>() where TProcessor : IPreProcessor;
        IPipelineBuilder UseConverter<TConverter>() where TConverter : IFileConverter;
        IPipelineBuilder UsePostProcessor<TProcessor>() where TProcessor : IPostProcessor;
        IPipelineBuilder UseOutputGenerator<TGenerator>() where TGenerator : IOutputGenerator;
        IProcessingPipeline Build();
    }

    /// <summary>
    /// Interface for executing processing pipelines
    /// </summary>
    public interface IProcessingPipeline
    {
        Task<ProcessingResult> ExecuteAsync(ProcessingContext context);
    }

    /// <summary>
    /// Context object that flows through the pipeline
    /// </summary>
    public class ProcessingContext
    {
        public string CorrelationId { get; }
        public IReadOnlyList<FileItem> InputFiles { get; }
        public Dictionary<string, object> Properties { get; }
        public CancellationToken CancellationToken { get; }
        public IProgress<ProcessingProgress> Progress { get; }
        public string OutputDirectory { get; set; }
        public ProcessingMode Mode { get; set; }

        public ProcessingContext(
            string correlationId,
            IReadOnlyList<FileItem> inputFiles,
            string outputDirectory,
            ProcessingMode mode,
            CancellationToken cancellationToken = default,
            IProgress<ProcessingProgress> progress = null)
        {
            CorrelationId = correlationId ?? Guid.NewGuid().ToString();
            InputFiles = inputFiles ?? throw new ArgumentNullException(nameof(inputFiles));
            OutputDirectory = outputDirectory ?? throw new ArgumentNullException(nameof(outputDirectory));
            Mode = mode;
            Properties = new Dictionary<string, object>();
            CancellationToken = cancellationToken;
            Progress = progress;
        }
    }

    /// <summary>
    /// Processing mode enumeration
    /// </summary>
    public enum ProcessingMode
    {
        SaveQuotes,
        DocumentUnlock,
        BulkConversion,
        FilePreview
    }

    /// <summary>
    /// Progress information for pipeline execution
    /// </summary>
    public class ProcessingProgress
    {
        public string CurrentStage { get; set; }
        public int CompletedFiles { get; set; }
        public int TotalFiles { get; set; }
        public double PercentComplete { get; set; }
        public string CurrentFileName { get; set; }
        public TimeSpan ElapsedTime { get; set; }
        public TimeSpan EstimatedRemaining { get; set; }
        public string StatusMessage { get; set; }
    }

    /// <summary>
    /// Result from pipeline execution
    /// </summary>
    public class ProcessingResult
    {
        public bool Success { get; set; }
        public List<string> SuccessfulFiles { get; set; } = new();
        public List<string> FailedFiles { get; set; } = new();
        public List<string> ErrorMessages { get; set; } = new();
        public Dictionary<string, object> OutputData { get; set; } = new();
        public TimeSpan ProcessingTime { get; set; }
        public bool IsMerged { get; set; }
    }
} 