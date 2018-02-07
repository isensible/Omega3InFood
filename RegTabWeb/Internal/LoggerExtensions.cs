using System;
using Microsoft.Extensions.Logging;

namespace RegTabWeb.Internal
{
    public static class LoggerExtensions
    {
        private static readonly Action<ILogger, Exception> _indexPageRequested;
        
        private static readonly Action<ILogger, string, Exception> _stataLogFileUploadRequested;
        
        private static readonly Action<ILogger, string, string, Exception> _excelDownloadRequested;
        
        private static readonly Action<ILogger, string, Exception> _excelDownloaded;
        
        private static readonly Action<ILogger, Exception> _logFileUploadFailed;
        private static readonly Action<ILogger, Exception> _excelDownloadFailed;
        
        static LoggerExtensions()
        {
            _indexPageRequested = LoggerMessage.Define(
                LogLevel.Information, 
                new EventId(1, nameof(IndexPageRequested)), 
                "GET request for Index page");

            _stataLogFileUploadRequested = LoggerMessage.Define<string>(
                LogLevel.Information,
                new EventId(2, nameof(StataLogFileUploadRequested)),
                "POST request for upload of stata log file {LogFileName}");
            
            _excelDownloadRequested = LoggerMessage.Define<string, string>(
                LogLevel.Information,
                new EventId(2, nameof(ExcelDownloadRequested)),
                "Downloading Excel generated from {LogFileName} with content {LogFileContent}");
            
            _excelDownloaded = LoggerMessage.Define<string>(
                LogLevel.Information,
                new EventId(2, nameof(ExcelDownloaded)),
                "Completed download of Excel generated from {LogFileName}");

            _logFileUploadFailed = LoggerMessage.Define(
                LogLevel.Error,
                new EventId(6, nameof(LogFileUploadFailed)),
                "Error uploading log file");
            
            _excelDownloadFailed = LoggerMessage.Define(
                LogLevel.Error,
                new EventId(6, nameof(ExcelDownloadFailed)),
                "Error downloading Excel file");
        }
        
        public static void IndexPageRequested(this ILogger logger)
        {
            _indexPageRequested(logger, null);
        }

        public static void StataLogFileUploadRequested(this ILogger logger, string logFileName)
        {
            _stataLogFileUploadRequested(logger, logFileName, null);
        }

        public static void ExcelDownloadRequested(this ILogger logger, string logFileName, string logFileContent)
        {
            _excelDownloadRequested(logger, logFileName, logFileContent, null);
        }

        public static void ExcelDownloaded(this ILogger logger, string logFileName)
        {
            _excelDownloaded(logger, logFileName, null);
        }

        public static void LogFileUploadFailed(this ILogger logger, Exception error)
        {
            _logFileUploadFailed(logger, error);
        }

        public static void ExcelDownloadFailed(this ILogger logger, Exception error)
        {
            _excelDownloadFailed(logger, error);
        }
    }
}