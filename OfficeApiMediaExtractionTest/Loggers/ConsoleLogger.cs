using OfficeApiMediaExtractionTest.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Loggers
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
