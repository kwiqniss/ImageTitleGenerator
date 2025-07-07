using ImageAnalyzer.Interfaces;
using System.Diagnostics;

namespace ImageAnalyzer.Loggers
{
    public class DebugLogger : ILogger
    {
        public void WriteLogMessage(string message)
        {
            if (!string.IsNullOrWhiteSpace(message))
                Debug.WriteLine(message);
        }
        public void WriteLogError(string? message, Exception? ex = null)
        {
            if (!string.IsNullOrWhiteSpace(message))
                Debug.WriteLine($"Error: {message}");

            if (ex != null)
                Debug.WriteLine($"Exception: {ex}");
        }
    }
}
