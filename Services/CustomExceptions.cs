using System;
using System.Runtime.InteropServices;

namespace DocHandler.Services
{
    /// <summary>
    /// Base exception for all DocHandler-specific errors with recovery guidance
    /// </summary>
    public abstract class DocHandlerException : Exception
    {
        public string RecoveryGuidance { get; }
        public string UserFriendlyMessage { get; }
        public ErrorSeverity Severity { get; }
        
        protected DocHandlerException(string message, string userFriendlyMessage, string recoveryGuidance, ErrorSeverity severity = ErrorSeverity.Error, Exception? innerException = null) 
            : base(message, innerException)
        {
            UserFriendlyMessage = userFriendlyMessage;
            RecoveryGuidance = recoveryGuidance;
            Severity = severity;
        }
    }

    public enum ErrorSeverity
    {
        Information,
        Warning,
        Error,
        Critical
    }

    #region Office/COM Related Exceptions

    /// <summary>
    /// Exception for Office/COM failures with specific recovery steps
    /// </summary>
    public class OfficeOperationException : DocHandlerException
    {
        public string OfficeApplication { get; }
        public int? ComErrorCode { get; }
        public bool IsRecoverable { get; }

        public OfficeOperationException(string officeApp, string operation, string userMessage, string recovery, bool recoverable = true, Exception? innerException = null, int? comErrorCode = null)
            : base($"Office {officeApp} operation '{operation}' failed", userMessage, recovery, ErrorSeverity.Error, innerException)
        {
            OfficeApplication = officeApp;
            IsRecoverable = recoverable;
            ComErrorCode = comErrorCode;
        }

        public static OfficeOperationException FromCOMException(string officeApp, string operation, COMException comEx)
        {
            var (userMessage, recovery, recoverable) = GetCOMErrorDetails(comEx.HResult);
            return new OfficeOperationException(officeApp, operation, userMessage, recovery, recoverable, comEx, comEx.HResult);
        }

        private static (string userMessage, string recovery, bool recoverable) GetCOMErrorDetails(int hResult)
        {
            return hResult switch
            {
                unchecked((int)0x800706BA) => (
                    "Microsoft Office is temporarily unavailable.", 
                    "Please wait a moment and try again. If the problem persists, restart the application.",
                    true),
                unchecked((int)0x80010001) => (
                    "Microsoft Office is busy with another operation.", 
                    "Please wait for the current operation to complete and try again.",
                    true),
                unchecked((int)0x80010105) => (
                    "Microsoft Office encountered an internal error.", 
                    "Try restarting Microsoft Office or the application.",
                    true),
                unchecked((int)0x8001010A) => (
                    "Microsoft Office is busy and cannot process the request.", 
                    "Please close any open Office documents and try again.",
                    true),
                unchecked((int)0x80004005) => (
                    "An unexpected error occurred in Microsoft Office.", 
                    "Try restarting the application. If the problem persists, repair Microsoft Office.",
                    true),
                unchecked((int)0x80070006) => (
                    "Microsoft Office file handle is invalid.", 
                    "Close the file in Office and try again.",
                    true),
                _ => (
                    "Microsoft Office encountered an unknown error.", 
                    "Try restarting the application. If the problem persists, contact support.",
                    false)
            };
        }
    }

    /// <summary>
    /// Exception for when Office applications crash or become unresponsive
    /// </summary>
    public class OfficeCrashException : DocHandlerException
    {
        public string OfficeApplication { get; }
        public int? ProcessId { get; }

        public OfficeCrashException(string officeApp, int? processId = null, Exception? innerException = null)
            : base($"Office application {officeApp} has crashed or become unresponsive", 
                   $"Microsoft {officeApp} has stopped responding.", 
                   "The application will attempt to restart Office automatically. Please try your operation again.",
                   ErrorSeverity.Critical, innerException)
        {
            OfficeApplication = officeApp;
            ProcessId = processId;
        }
    }

    #endregion

    #region File Processing Exceptions

    /// <summary>
    /// Exception for file validation failures with specific guidance
    /// </summary>
    public class FileValidationException : DocHandlerException
    {
        public string FilePath { get; }
        public ValidationFailureReason Reason { get; }

        public FileValidationException(string filePath, ValidationFailureReason reason, string details = "", Exception? innerException = null)
            : base($"File validation failed: {filePath}", GetUserMessage(reason, filePath), GetRecoveryGuidance(reason), GetSeverity(reason), innerException)
        {
            FilePath = filePath;
            Reason = reason;
        }

        private static string GetUserMessage(ValidationFailureReason reason, string filePath)
        {
            var fileName = System.IO.Path.GetFileName(filePath);
            return reason switch
            {
                ValidationFailureReason.FileNotFound => $"The file '{fileName}' could not be found.",
                ValidationFailureReason.FileTooLarge => $"The file '{fileName}' is too large to process.",
                ValidationFailureReason.UnsupportedFileType => $"The file type of '{fileName}' is not supported.",
                ValidationFailureReason.FileCorrupted => $"The file '{fileName}' appears to be corrupted or damaged.",
                ValidationFailureReason.AccessDenied => $"Access to the file '{fileName}' was denied.",
                ValidationFailureReason.FileLocked => $"The file '{fileName}' is currently in use by another application.",
                ValidationFailureReason.SecurityViolation => $"The file '{fileName}' failed security validation.",
                ValidationFailureReason.PathTraversal => $"The file path contains invalid or unsafe characters.",
                ValidationFailureReason.EmptyFile => $"The file '{fileName}' is empty.",
                _ => $"The file '{fileName}' failed validation."
            };
        }

        private static string GetRecoveryGuidance(ValidationFailureReason reason)
        {
            return reason switch
            {
                ValidationFailureReason.FileNotFound => "Please check that the file exists and try again.",
                ValidationFailureReason.FileTooLarge => "Try using a smaller file (maximum size is 50MB).",
                ValidationFailureReason.UnsupportedFileType => "Please use a supported file type (.pdf, .docx, .doc, .xlsx, .xls).",
                ValidationFailureReason.FileCorrupted => "Try opening the file in its native application to check if it can be repaired.",
                ValidationFailureReason.AccessDenied => "Check file permissions or run the application as administrator.",
                ValidationFailureReason.FileLocked => "Close the file in other applications and try again.",
                ValidationFailureReason.SecurityViolation => "Use a different file or contact your administrator.",
                ValidationFailureReason.PathTraversal => "Use a simpler file path without special characters.",
                ValidationFailureReason.EmptyFile => "Use a file that contains content.",
                _ => "Please try with a different file."
            };
        }

        private static ErrorSeverity GetSeverity(ValidationFailureReason reason)
        {
            return reason switch
            {
                ValidationFailureReason.SecurityViolation => ErrorSeverity.Critical,
                ValidationFailureReason.PathTraversal => ErrorSeverity.Critical,
                ValidationFailureReason.FileCorrupted => ErrorSeverity.Error,
                ValidationFailureReason.UnsupportedFileType => ErrorSeverity.Warning,
                _ => ErrorSeverity.Error
            };
        }
    }

    public enum ValidationFailureReason
    {
        FileNotFound,
        FileTooLarge,
        UnsupportedFileType,
        FileCorrupted,
        AccessDenied,
        FileLocked,
        SecurityViolation,
        PathTraversal,
        EmptyFile,
        Unknown
    }

    /// <summary>
    /// Exception for file processing operations
    /// </summary>
    public class FileProcessingException : DocHandlerException
    {
        public string FilePath { get; }
        public string Operation { get; }

        public FileProcessingException(string filePath, string operation, string userMessage, string recovery, Exception? innerException = null)
            : base($"File processing failed: {operation} on {filePath}", userMessage, recovery, ErrorSeverity.Error, innerException)
        {
            FilePath = filePath;
            Operation = operation;
        }
    }

    #endregion

    #region Security Exceptions

    /// <summary>
    /// Exception for security violations requiring immediate attention
    /// </summary>
    public class SecurityViolationException : DocHandlerException
    {
        public string ViolationType { get; }
        public string Resource { get; }

        public SecurityViolationException(string violationType, string resource, string details, Exception? innerException = null)
            : base($"Security violation: {violationType} on {resource}", 
                   "A security violation was detected.", 
                   "The operation has been blocked for security reasons. Contact your administrator if you believe this is an error.",
                   ErrorSeverity.Critical, innerException)
        {
            ViolationType = violationType;
            Resource = resource;
        }
    }

    #endregion

    #region Configuration and Service Exceptions

    /// <summary>
    /// Exception for configuration-related errors
    /// </summary>
    public class ConfigurationException : DocHandlerException
    {
        public string ConfigurationKey { get; }

        public ConfigurationException(string configKey, string userMessage, string recovery, Exception? innerException = null)
            : base($"Configuration error: {configKey}", userMessage, recovery, ErrorSeverity.Error, innerException)
        {
            ConfigurationKey = configKey;
        }
    }

    /// <summary>
    /// Exception for service initialization or operation failures
    /// </summary>
    public class ServiceException : DocHandlerException
    {
        public string ServiceName { get; }
        public string Operation { get; }

        public ServiceException(string serviceName, string operation, string userMessage, string recovery, Exception? innerException = null)
            : base($"Service error: {serviceName} - {operation}", userMessage, recovery, ErrorSeverity.Error, innerException)
        {
            ServiceName = serviceName;
            Operation = operation;
        }
    }

    #endregion

    #region Utility Classes

    /// <summary>
    /// Helper class for creating common exceptions
    /// </summary>
    public static class ExceptionFactory
    {
        public static FileValidationException FileNotFound(string filePath)
        {
            return new FileValidationException(filePath, ValidationFailureReason.FileNotFound);
        }

        public static FileValidationException FileTooLarge(string filePath, long actualSize, long maxSize)
        {
            return new FileValidationException(filePath, ValidationFailureReason.FileTooLarge, 
                $"File size: {FormatFileSize(actualSize)}, Maximum: {FormatFileSize(maxSize)}");
        }

        public static FileValidationException UnsupportedFileType(string filePath, string extension)
        {
            return new FileValidationException(filePath, ValidationFailureReason.UnsupportedFileType, 
                $"Extension: {extension}");
        }

        public static OfficeOperationException OfficeNotAvailable(string officeApp)
        {
            return new OfficeOperationException(officeApp, "Initialize", 
                $"Microsoft {officeApp} is not available on this system.", 
                $"Please install Microsoft {officeApp} to use this feature.", false);
        }

        public static SecurityViolationException PathTraversal(string path)
        {
            return new SecurityViolationException("Path Traversal", path, 
                "File path contains potentially unsafe characters or traversal attempts.");
        }

        private static string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }
    }

    #endregion
} 