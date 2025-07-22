using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Serilog;

namespace DocHandler.Services
{
    /// <summary>
    /// Comprehensive error recovery service that handles exceptions and provides automatic recovery
    /// </summary>
    public class ErrorRecoveryService : IDisposable
    {
        private readonly ILogger _logger;
        private readonly ProcessManager? _processManager;
        private readonly ConfigurationService? _configService;
        private readonly Dictionary<Type, Func<Exception, Task<RecoveryResult>>> _recoveryStrategies;
        private readonly SemaphoreSlim _recoveryLock = new(1, 1);
        private bool _disposed;

        // Recovery statistics
        private int _totalRecoveryAttempts;
        private int _successfulRecoveries;
        private DateTime _lastRecoveryTime = DateTime.MinValue;

        public ErrorRecoveryService(ProcessManager? processManager = null, ConfigurationService? configService = null)
        {
            _logger = Log.ForContext<ErrorRecoveryService>();
            _processManager = processManager;
            _configService = configService;
            _recoveryStrategies = InitializeRecoveryStrategies();
            
            _logger.Information("Error recovery service initialized with {StrategyCount} recovery strategies", 
                _recoveryStrategies.Count);
        }

        /// <summary>
        /// Handles an exception and attempts recovery based on exception type
        /// </summary>
        public async Task<RecoveryResult> HandleExceptionAsync(Exception exception, string context = "")
        {
            if (_disposed) return RecoveryResult.CreateFailed("Service disposed");

            await _recoveryLock.WaitAsync();
            try
            {
                _totalRecoveryAttempts++;
                _lastRecoveryTime = DateTime.UtcNow;

                _logger.Error(exception, "Handling exception in context: {Context}", context);

                // Try specific recovery strategy first
                var specificResult = await TrySpecificRecoveryAsync(exception);
                if (specificResult.Success)
                {
                    _successfulRecoveries++;
                    return specificResult;
                }

                // Fall back to generic recovery
                var genericResult = await TryGenericRecoveryAsync(exception, context);
                if (genericResult.Success)
                {
                    _successfulRecoveries++;
                }

                return genericResult;
            }
            finally
            {
                _recoveryLock.Release();
            }
        }

        /// <summary>
        /// Creates user-friendly error information from any exception
        /// </summary>
        public ErrorInfo CreateErrorInfo(Exception exception, string context = "")
        {
            return exception switch
            {
                DocHandlerException dhEx => new ErrorInfo
                {
                    Title = GetErrorTitle(dhEx.Severity),
                    Message = dhEx.UserFriendlyMessage,
                    Details = dhEx.Message,
                    RecoveryGuidance = dhEx.RecoveryGuidance,
                    Severity = dhEx.Severity,
                    CanRetry = dhEx is OfficeOperationException officeEx && officeEx.IsRecoverable,
                    Context = context
                },
                
                COMException comEx => new ErrorInfo
                {
                    Title = "Microsoft Office Error",
                    Message = GetCOMErrorMessage(comEx.HResult),
                    Details = comEx.Message,
                    RecoveryGuidance = GetCOMRecoveryGuidance(comEx.HResult),
                    Severity = ErrorSeverity.Error,
                    CanRetry = IsRecoverableCOMError(comEx.HResult),
                    Context = context
                },
                
                FileNotFoundException => new ErrorInfo
                {
                    Title = "File Not Found",
                    Message = "The specified file could not be found.",
                    Details = exception.Message,
                    RecoveryGuidance = "Please check that the file exists and you have permission to access it.",
                    Severity = ErrorSeverity.Error,
                    CanRetry = true,
                    Context = context
                },
                
                UnauthorizedAccessException => new ErrorInfo
                {
                    Title = "Access Denied",
                    Message = "Access to the file or resource was denied.",
                    Details = exception.Message,
                    RecoveryGuidance = "Check file permissions or run the application as administrator.",
                    Severity = ErrorSeverity.Error,
                    CanRetry = true,
                    Context = context
                },
                
                IOException ioEx => new ErrorInfo
                {
                    Title = "File Operation Error",
                    Message = GetIOErrorMessage(ioEx),
                    Details = ioEx.Message,
                    RecoveryGuidance = GetIORecoveryGuidance(ioEx),
                    Severity = ErrorSeverity.Error,
                    CanRetry = true,
                    Context = context
                },
                
                TimeoutException => new ErrorInfo
                {
                    Title = "Operation Timed Out",
                    Message = "The operation took too long to complete.",
                    Details = exception.Message,
                    RecoveryGuidance = "Try again with a smaller file or increase the timeout setting.",
                    Severity = ErrorSeverity.Warning,
                    CanRetry = true,
                    Context = context
                },
                
                _ => new ErrorInfo
                {
                    Title = "Unexpected Error",
                    Message = "An unexpected error occurred.",
                    Details = exception.Message,
                    RecoveryGuidance = "Please try again. If the problem persists, contact support.",
                    Severity = ErrorSeverity.Error,
                    CanRetry = true,
                    Context = context
                }
            };
        }

        /// <summary>
        /// Shows an error dialog to the user with recovery options
        /// </summary>
        public async Task<UserRecoveryChoice> ShowErrorDialogAsync(Exception exception, string context = "")
        {
            var errorInfo = CreateErrorInfo(exception, context);
            
            return await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var message = $"{errorInfo.Message}\n\n{errorInfo.RecoveryGuidance}";
                
                if (errorInfo.CanRetry)
                {
                    var result = MessageBox.Show(
                        $"{message}\n\nWould you like to try again?",
                        errorInfo.Title,
                        MessageBoxButton.YesNo,
                        GetMessageBoxImage(errorInfo.Severity));
                    
                    return result == MessageBoxResult.Yes ? UserRecoveryChoice.Retry : UserRecoveryChoice.Cancel;
                }
                else
                {
                    MessageBox.Show(message, errorInfo.Title, MessageBoxButton.OK, GetMessageBoxImage(errorInfo.Severity));
                    return UserRecoveryChoice.Acknowledge;
                }
            });
        }

        /// <summary>
        /// Gets recovery statistics
        /// </summary>
        public RecoveryStatistics GetStatistics()
        {
            return new RecoveryStatistics
            {
                TotalRecoveryAttempts = _totalRecoveryAttempts,
                SuccessfulRecoveries = _successfulRecoveries,
                SuccessRate = _totalRecoveryAttempts > 0 ? (double)_successfulRecoveries / _totalRecoveryAttempts : 0.0,
                LastRecoveryTime = _lastRecoveryTime
            };
        }

        #region Private Methods

        private Dictionary<Type, Func<Exception, Task<RecoveryResult>>> InitializeRecoveryStrategies()
        {
            return new Dictionary<Type, Func<Exception, Task<RecoveryResult>>>
            {
                { typeof(OfficeOperationException), RecoverOfficeOperationAsync },
                { typeof(OfficeCrashException), RecoverOfficeCrashAsync },
                { typeof(FileValidationException), RecoverFileValidationAsync },
                { typeof(FileProcessingException), RecoverFileProcessingAsync },
                { typeof(SecurityViolationException), RecoverSecurityViolationAsync },
                { typeof(COMException), RecoverCOMExceptionAsync }
            };
        }

        private async Task<RecoveryResult> TrySpecificRecoveryAsync(Exception exception)
        {
            var exceptionType = exception.GetType();
            
            // Try exact type match first
            if (_recoveryStrategies.TryGetValue(exceptionType, out var strategy))
            {
                try
                {
                    return await strategy(exception);
                }
                catch (Exception recoveryEx)
                {
                    _logger.Warning(recoveryEx, "Recovery strategy failed for {ExceptionType}", exceptionType.Name);
                }
            }
            
            // Try base type matches
            foreach (var (strategyType, strategyFunc) in _recoveryStrategies)
            {
                if (strategyType.IsAssignableFrom(exceptionType))
                {
                    try
                    {
                        return await strategyFunc(exception);
                    }
                    catch (Exception recoveryEx)
                    {
                        _logger.Warning(recoveryEx, "Base type recovery strategy failed for {ExceptionType}", strategyType.Name);
                    }
                }
            }

            return RecoveryResult.CreateFailed("No specific recovery strategy found");
        }

        private async Task<RecoveryResult> TryGenericRecoveryAsync(Exception exception, string context)
        {
            _logger.Information("Attempting generic recovery for {ExceptionType} in context {Context}", 
                exception.GetType().Name, context);

            // Generic cleanup and reset
            try
            {
                // Force garbage collection to free resources
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Small delay to allow system recovery
                await Task.Delay(1000);

                return RecoveryResult.CreateSuccess("Generic recovery completed");
            }
            catch (Exception recoveryEx)
            {
                _logger.Error(recoveryEx, "Generic recovery failed");
                return RecoveryResult.CreateFailed($"Generic recovery failed: {recoveryEx.Message}");
            }
        }

        private async Task<RecoveryResult> RecoverOfficeOperationAsync(Exception exception)
        {
            var officeEx = (OfficeOperationException)exception;
            _logger.Information("Attempting Office operation recovery for {OfficeApp}", officeEx.OfficeApplication);

            if (!officeEx.IsRecoverable)
            {
                return RecoveryResult.CreateFailed("Office operation is not recoverable");
            }

            try
            {
                // Kill any hanging Office processes
                if (_processManager != null)
                {
                    if (officeEx.OfficeApplication.Contains("Word"))
                    {
                        _processManager.KillOrphanedWordProcesses();
                    }
                    else if (officeEx.OfficeApplication.Contains("Excel"))
                    {
                        _processManager.KillOrphanedExcelProcesses();
                    }
                }

                // Wait for processes to clean up
                await Task.Delay(2000);

                // Force COM cleanup
                ComHelper.ForceComCleanup();

                return RecoveryResult.CreateSuccess("Office processes cleaned up successfully");
            }
            catch (Exception recoveryEx)
            {
                _logger.Error(recoveryEx, "Office recovery failed");
                return RecoveryResult.CreateFailed($"Office recovery failed: {recoveryEx.Message}");
            }
        }

        private async Task<RecoveryResult> RecoverOfficeCrashAsync(Exception exception)
        {
            var crashEx = (OfficeCrashException)exception;
            _logger.Fatal("Recovering from Office crash: {OfficeApp}", crashEx.OfficeApplication);

            try
            {
                // Immediate process cleanup
                if (_processManager != null)
                {
                    _processManager.TerminateOrphanedOfficeProcesses();
                }

                // Clear COM state
                ComHelper.ForceComCleanup();

                // Wait for cleanup
                await Task.Delay(3000);

                return RecoveryResult.CreateSuccess("Office crash recovery completed");
            }
            catch (Exception recoveryEx)
            {
                _logger.Error(recoveryEx, "Office crash recovery failed");
                return RecoveryResult.CreateFailed($"Crash recovery failed: {recoveryEx.Message}");
            }
        }

        private async Task<RecoveryResult> RecoverFileValidationAsync(Exception exception)
        {
            var fileEx = (FileValidationException)exception;
            _logger.Information("Attempting file validation recovery for {Reason}", fileEx.Reason);

            // Most file validation errors can't be automatically recovered
            // But we can provide specific guidance
            return fileEx.Reason switch
            {
                ValidationFailureReason.FileLocked => await RetryAfterDelay(1000),
                ValidationFailureReason.AccessDenied => RecoveryResult.CreateFailed("Access denied - manual intervention required"),
                ValidationFailureReason.SecurityViolation => RecoveryResult.CreateFailed("Security violation - operation blocked"),
                _ => RecoveryResult.CreateFailed("File validation error - manual intervention required")
            };
        }

        private async Task<RecoveryResult> RecoverFileProcessingAsync(Exception exception)
        {
            var fileEx = (FileProcessingException)exception;
            _logger.Information("Attempting file processing recovery for operation {Operation}", fileEx.Operation);

            try
            {
                // Generic file processing recovery
                await Task.Delay(500);
                return RecoveryResult.CreateSuccess("File processing recovery completed");
            }
            catch (Exception recoveryEx)
            {
                return RecoveryResult.CreateFailed($"File processing recovery failed: {recoveryEx.Message}");
            }
        }

        private async Task<RecoveryResult> RecoverSecurityViolationAsync(Exception exception)
        {
            var secEx = (SecurityViolationException)exception;
            _logger.Fatal("Security violation detected: {ViolationType} on {Resource}", 
                secEx.ViolationType, secEx.Resource);

            // Security violations should not be automatically recovered
            return RecoveryResult.CreateFailed("Security violation - operation permanently blocked");
        }

        private async Task<RecoveryResult> RecoverCOMExceptionAsync(Exception exception)
        {
            var comEx = (COMException)exception;
            _logger.Information("Attempting COM exception recovery for HRESULT {HResult:X8}", comEx.HResult);

            if (!IsRecoverableCOMError(comEx.HResult))
            {
                return RecoveryResult.CreateFailed("COM error is not recoverable");
            }

            try
            {
                // Standard COM recovery
                ComHelper.ForceComCleanup();
                await Task.Delay(1000);
                return RecoveryResult.CreateSuccess("COM exception recovery completed");
            }
            catch (Exception recoveryEx)
            {
                return RecoveryResult.CreateFailed($"COM recovery failed: {recoveryEx.Message}");
            }
        }

        private async Task<RecoveryResult> RetryAfterDelay(int delayMs)
        {
            await Task.Delay(delayMs);
            return RecoveryResult.CreateSuccess($"Delayed retry after {delayMs}ms");
        }

        private static string GetErrorTitle(ErrorSeverity severity)
        {
            return severity switch
            {
                ErrorSeverity.Information => "Information",
                ErrorSeverity.Warning => "Warning",
                ErrorSeverity.Error => "Error",
                ErrorSeverity.Critical => "Critical Error",
                _ => "Error"
            };
        }

        private static MessageBoxImage GetMessageBoxImage(ErrorSeverity severity)
        {
            return severity switch
            {
                ErrorSeverity.Information => MessageBoxImage.Information,
                ErrorSeverity.Warning => MessageBoxImage.Warning,
                ErrorSeverity.Error => MessageBoxImage.Error,
                ErrorSeverity.Critical => MessageBoxImage.Error,
                _ => MessageBoxImage.Error
            };
        }

        private static string GetCOMErrorMessage(int hResult)
        {
            return hResult switch
            {
                unchecked((int)0x800706BA) => "Microsoft Office is temporarily unavailable.",
                unchecked((int)0x80010001) => "Microsoft Office is busy with another operation.",
                unchecked((int)0x80010105) => "Microsoft Office encountered an internal error.",
                unchecked((int)0x8001010A) => "Microsoft Office is busy and cannot process the request.",
                _ => "Microsoft Office encountered an error."
            };
        }

        private static string GetCOMRecoveryGuidance(int hResult)
        {
            return hResult switch
            {
                unchecked((int)0x800706BA) => "Please wait a moment and try again.",
                unchecked((int)0x80010001) => "Please wait for the current operation to complete.",
                unchecked((int)0x80010105) => "Try restarting Microsoft Office.",
                unchecked((int)0x8001010A) => "Close any open Office documents and try again.",
                _ => "Try restarting the application."
            };
        }

        private static bool IsRecoverableCOMError(int hResult)
        {
            var recoverableErrors = new[]
            {
                unchecked((int)0x800706BA), // RPC server unavailable
                unchecked((int)0x80010001), // Call was rejected by callee
                unchecked((int)0x80010105), // Server threw an exception
                unchecked((int)0x8001010A), // Message filter indicated application is busy
                unchecked((int)0x80004005), // Unspecified error (sometimes recoverable)
            };

            return recoverableErrors.Contains(hResult);
        }

        private static string GetIOErrorMessage(IOException ioEx)
        {
            var message = ioEx.Message?.ToLowerInvariant() ?? "";
            
            if (message.Contains("being used by another process"))
                return "The file is currently open in another application.";
            if (message.Contains("access denied"))
                return "Access to the file was denied.";
            if (message.Contains("sharing violation"))
                return "The file cannot be accessed because it is being used by another process.";
            
            return "A file operation error occurred.";
        }

        private static string GetIORecoveryGuidance(IOException ioEx)
        {
            var message = ioEx.Message?.ToLowerInvariant() ?? "";
            
            if (message.Contains("being used by another process") || message.Contains("sharing violation"))
                return "Close the file in other applications and try again.";
            if (message.Contains("access denied"))
                return "Check file permissions or run as administrator.";
            
            return "Please try again or use a different file.";
        }

        #endregion

        public void Dispose()
        {
            if (!_disposed)
            {
                _recoveryLock?.Dispose();
                _disposed = true;
                _logger.Information("Error recovery service disposed");
            }
        }
    }

    #region Supporting Classes

    public class RecoveryResult
    {
        public bool Success { get; private set; }
        public string Message { get; private set; }
        public TimeSpan Duration { get; private set; }

        private RecoveryResult(bool success, string message)
        {
            Success = success;
            Message = message;
            Duration = TimeSpan.Zero;
        }

        public static RecoveryResult CreateSuccess(string message) => new(true, message);
        public static RecoveryResult CreateFailed(string message) => new(false, message);
    }

    public class ErrorInfo
    {
        public string Title { get; set; } = "";
        public string Message { get; set; } = "";
        public string Details { get; set; } = "";
        public string RecoveryGuidance { get; set; } = "";
        public ErrorSeverity Severity { get; set; }
        public bool CanRetry { get; set; }
        public string Context { get; set; } = "";
    }

    public enum UserRecoveryChoice
    {
        Retry,
        Cancel,
        Acknowledge
    }

    public class RecoveryStatistics
    {
        public int TotalRecoveryAttempts { get; set; }
        public int SuccessfulRecoveries { get; set; }
        public double SuccessRate { get; set; }
        public DateTime LastRecoveryTime { get; set; }
    }

    #endregion
} 