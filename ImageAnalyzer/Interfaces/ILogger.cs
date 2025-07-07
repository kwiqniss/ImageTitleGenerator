namespace ImageAnalyzer.Interfaces
{
    public interface ILogger
    {
        void WriteLogMessage(string message);
        void WriteLogError(string? message, Exception? ex = null);
    }
}
