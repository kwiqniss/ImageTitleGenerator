namespace ImageAnalyzerProgram.Loggers
{
    public interface ILogger
    {
        void WriteLogMessage(string message);
        void WriteLogError(string? message, Exception? ex = null);
    }
}
