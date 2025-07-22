using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using DocHandler.Models;
using Microsoft.Extensions.DependencyInjection;
using Serilog;

namespace DocHandler.Services.Pipeline
{
    /// <summary>
    /// Builder for creating processing pipelines
    /// </summary>
    public class PipelineBuilder : IPipelineBuilder
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly List<Type> _validators = new();
        private readonly List<Type> _preProcessors = new();
        private readonly List<Type> _converters = new();
        private readonly List<Type> _postProcessors = new();
        private readonly List<Type> _outputGenerators = new();

        public PipelineBuilder(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
        }

        public IPipelineBuilder UseValidator<TValidator>() where TValidator : IFileValidator
        {
            _validators.Add(typeof(TValidator));
            return this;
        }

        public IPipelineBuilder UsePreProcessor<TProcessor>() where TProcessor : IPreProcessor
        {
            _preProcessors.Add(typeof(TProcessor));
            return this;
        }

        public IPipelineBuilder UseConverter<TConverter>() where TConverter : IFileConverter
        {
            _converters.Add(typeof(TConverter));
            return this;
        }

        public IPipelineBuilder UsePostProcessor<TProcessor>() where TProcessor : IPostProcessor
        {
            _postProcessors.Add(typeof(TProcessor));
            return this;
        }

        public IPipelineBuilder UseOutputGenerator<TGenerator>() where TGenerator : IOutputGenerator
        {
            _outputGenerators.Add(typeof(TGenerator));
            return this;
        }

        public IProcessingPipeline Build()
        {
            return new ProcessingPipeline(
                _serviceProvider,
                _validators,
                _preProcessors,
                _converters,
                _postProcessors,
                _outputGenerators);
        }
    }

    /// <summary>
    /// Concrete implementation of processing pipeline
    /// </summary>
    public class ProcessingPipeline : IProcessingPipeline
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly List<Type> _validators;
        private readonly List<Type> _preProcessors;
        private readonly List<Type> _converters;
        private readonly List<Type> _postProcessors;
        private readonly List<Type> _outputGenerators;
        private readonly ILogger _logger;

        public ProcessingPipeline(
            IServiceProvider serviceProvider,
            List<Type> validators,
            List<Type> preProcessors,
            List<Type> converters,
            List<Type> postProcessors,
            List<Type> outputGenerators)
        {
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _validators = validators ?? new List<Type>();
            _preProcessors = preProcessors ?? new List<Type>();
            _converters = converters ?? new List<Type>();
            _postProcessors = postProcessors ?? new List<Type>();
            _outputGenerators = outputGenerators ?? new List<Type>();
            _logger = Log.ForContext<ProcessingPipeline>();
        }

        public async Task<ProcessingResult> ExecuteAsync(ProcessingContext context)
        {
            var stopwatch = Stopwatch.StartNew();
            var result = new ProcessingResult();
            
            try
            {
                _logger.Information("Starting pipeline execution for {FileCount} files", context.InputFiles.Count);

                // Step 1: Validation
                var validFiles = await ExecuteValidationStageAsync(context, result);
                if (!validFiles.Any())
                {
                    result.Success = false;
                    result.ErrorMessages.Add("No valid files to process");
                    return result;
                }

                // Step 2: Pre-processing
                var preProcessedFiles = await ExecutePreProcessingStageAsync(validFiles, context, result);

                // Step 3: Conversion
                var conversions = await ExecuteConversionStageAsync(preProcessedFiles, context, result);

                // Step 4: Post-processing
                var postProcessed = await ExecutePostProcessingStageAsync(conversions, context, result);

                // Step 5: Output generation
                await ExecuteOutputGenerationStageAsync(postProcessed, context, result);

                result.Success = result.SuccessfulFiles.Any();
                result.ProcessingTime = stopwatch.Elapsed;

                _logger.Information("Pipeline execution completed in {Duration}ms. Success: {Success}, Files: {SuccessCount}/{TotalCount}",
                    stopwatch.ElapsedMilliseconds, result.Success, result.SuccessfulFiles.Count, context.InputFiles.Count);

                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Pipeline execution failed");
                result.Success = false;
                result.ErrorMessages.Add($"Pipeline execution failed: {ex.Message}");
                result.ProcessingTime = stopwatch.Elapsed;
                return result;
            }
        }

        private async Task<List<FileItem>> ExecuteValidationStageAsync(ProcessingContext context, ProcessingResult result)
        {
            var validFiles = new List<FileItem>();

            foreach (var file in context.InputFiles)
            {
                var isValid = true;

                foreach (var validatorType in _validators)
                {
                    var validator = (IFileValidator)_serviceProvider.GetRequiredService(validatorType);
                    
                    if (!await validator.CanProcessAsync(file, context))
                        continue;

                    var validationResult = await validator.ValidateAsync(file, context);
                    if (!validationResult.IsValid)
                    {
                        isValid = false;
                        result.FailedFiles.Add(file.FilePath);
                        result.ErrorMessages.AddRange(validationResult.ErrorMessages);
                        break;
                    }
                }

                if (isValid)
                {
                    validFiles.Add(file);
                }
            }

            ReportProgress(context, "Validation", validFiles.Count, context.InputFiles.Count, "Validating files...");
            return validFiles;
        }

        private async Task<List<PreProcessingResult>> ExecutePreProcessingStageAsync(List<FileItem> files, ProcessingContext context, ProcessingResult result)
        {
            var results = new List<PreProcessingResult>();

            for (int i = 0; i < files.Count; i++)
            {
                var file = files[i];
                var preProcessResult = new PreProcessingResult { ProcessedFile = file, Success = true };

                foreach (var processorType in _preProcessors)
                {
                    var processor = (IPreProcessor)_serviceProvider.GetRequiredService(processorType);
                    
                    if (!await processor.CanProcessAsync(file, context))
                        continue;

                    try
                    {
                        var stageResult = await processor.ProcessAsync(file, context);
                        if (!stageResult.Success)
                        {
                            preProcessResult.Success = false;
                            preProcessResult.Error = stageResult.Error;
                            break;
                        }

                        // Merge extracted data
                        foreach (var kvp in stageResult.ExtractedData)
                        {
                            preProcessResult.ExtractedData[kvp.Key] = kvp.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Pre-processing failed for file {FilePath}", file.FilePath);
                        preProcessResult.Success = false;
                        preProcessResult.Error = ex;
                        break;
                    }
                }

                results.Add(preProcessResult);
                ReportProgress(context, "Pre-processing", i + 1, files.Count, $"Pre-processing {file.FileName}...");
            }

            return results;
        }

        private async Task<List<ConversionResult>> ExecuteConversionStageAsync(List<PreProcessingResult> preProcessed, ProcessingContext context, ProcessingResult result)
        {
            var conversions = new List<ConversionResult>();

            for (int i = 0; i < preProcessed.Count; i++)
            {
                var preProcess = preProcessed[i];
                if (!preProcess.Success)
                {
                    result.FailedFiles.Add(preProcess.ProcessedFile.FilePath);
                    continue;
                }

                ConversionResult conversionResult = null;

                foreach (var converterType in _converters)
                {
                    var converter = (IFileConverter)_serviceProvider.GetRequiredService(converterType);
                    
                    if (!await converter.CanProcessAsync(preProcess.ProcessedFile, context))
                        continue;

                    try
                    {
                        conversionResult = await converter.ConvertAsync(preProcess.ProcessedFile, context);
                        if (conversionResult.Success)
                        {
                            break; // Success, use this conversion
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Conversion failed for file {FilePath}", preProcess.ProcessedFile.FilePath);
                        conversionResult = new ConversionResult
                        {
                            Success = false,
                            SourceFile = preProcess.ProcessedFile,
                            Error = ex
                        };
                    }
                }

                if (conversionResult != null)
                {
                    conversions.Add(conversionResult);

                    if (conversionResult.Success)
                    {
                        result.SuccessfulFiles.Add(conversionResult.OutputPath);
                    }
                    else
                    {
                        result.FailedFiles.Add(preProcess.ProcessedFile.FilePath);
                        result.ErrorMessages.Add($"Conversion failed for {preProcess.ProcessedFile.FileName}");
                    }
                }

                ReportProgress(context, "Conversion", i + 1, preProcessed.Count, $"Converting {preProcess.ProcessedFile.FileName}...");
            }

            return conversions;
        }

        private async Task<List<PostProcessingResult>> ExecutePostProcessingStageAsync(List<ConversionResult> conversions, ProcessingContext context, ProcessingResult result)
        {
            var postProcessed = new List<PostProcessingResult>();

            for (int i = 0; i < conversions.Count; i++)
            {
                var conversion = conversions[i];
                if (!conversion.Success)
                {
                    continue;
                }

                var postProcessResult = new PostProcessingResult
                {
                    Success = true,
                    SourceConversion = conversion,
                    FinalPath = conversion.OutputPath
                };

                foreach (var processorType in _postProcessors)
                {
                    var processor = (IPostProcessor)_serviceProvider.GetRequiredService(processorType);
                    
                    if (!await processor.CanProcessAsync(conversion.SourceFile, context))
                        continue;

                    try
                    {
                        var stageResult = await processor.ProcessAsync(conversion, context);
                        if (!stageResult.Success)
                        {
                            postProcessResult.Success = false;
                            postProcessResult.Error = stageResult.Error;
                            break;
                        }

                        postProcessResult.FinalPath = stageResult.FinalPath;
                        // Merge post-processing data
                        foreach (var kvp in stageResult.PostProcessingData)
                        {
                            postProcessResult.PostProcessingData[kvp.Key] = kvp.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.Error(ex, "Post-processing failed for file {FilePath}", conversion.OutputPath);
                        postProcessResult.Success = false;
                        postProcessResult.Error = ex;
                        break;
                    }
                }

                postProcessed.Add(postProcessResult);
                ReportProgress(context, "Post-processing", i + 1, conversions.Count, $"Post-processing {conversion.SourceFile.FileName}...");
            }

            return postProcessed;
        }

        private async Task ExecuteOutputGenerationStageAsync(List<PostProcessingResult> postProcessed, ProcessingContext context, ProcessingResult result)
        {
            foreach (var generatorType in _outputGenerators)
            {
                var generator = (IOutputGenerator)_serviceProvider.GetRequiredService(generatorType);
                
                try
                {
                    var outputResult = await generator.GenerateAsync(postProcessed, context);
                    if (outputResult.Success)
                    {
                        // Merge output metadata
                        foreach (var kvp in outputResult.OutputMetadata)
                        {
                            result.OutputData[kvp.Key] = kvp.Value;
                        }
                    }
                    else
                    {
                        result.ErrorMessages.AddRange(outputResult.Messages);
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error(ex, "Output generation failed");
                    result.ErrorMessages.Add($"Output generation failed: {ex.Message}");
                }
            }

            ReportProgress(context, "Output Generation", 1, 1, "Finalizing output...");
        }

        private void ReportProgress(ProcessingContext context, string stage, int completed, int total, string message)
        {
            if (context.Progress == null) return;

            var progress = new ProcessingProgress
            {
                CurrentStage = stage,
                CompletedFiles = completed,
                TotalFiles = total,
                PercentComplete = total > 0 ? (double)completed / total * 100 : 0,
                StatusMessage = message
            };

            context.Progress.Report(progress);
        }
    }
} 