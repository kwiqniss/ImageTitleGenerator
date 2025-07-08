namespace ImageAnalyzerProgram.Loggers
{
    public class ConsoleLogger : ILogger
    {
        public void WriteLogMessage(string message)
        {
            Console.WriteLine(message);
        }
        public void WriteLogError(string? message, Exception? ex = null)
        {
            if (!string.IsNullOrWhiteSpace(message))
                Console.Error.WriteLine($"Error: {message}");
            if (ex != null)
                Console.Error.WriteLine($"Exception: {ex}");
        }
    }
}
