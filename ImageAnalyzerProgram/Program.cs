// See https://aka.ms/new-console-template for more information
using ImageAnalyzer.AI.Acs;
using ImageAnalyzer.DataTypes;
using ImageAnalyzer.Interfaces;
using ImageAnalyzer.IO;
using ImageAnalyzer.Office;
using ImageAnalyzerProgram.Loggers;

namespace ImageAnalyzerProgram
{
    public class Program
    {
        private static string _docPath = string.Empty;
        private static AcsConnectionInfo? _acsConnectionDetails;
        private static List<ILogger>? _loggers;
        private static IWorker? _worker;

        public static async Task Main(string[] args)
        {
            _loggers = new List<ILogger>
                {
                    new ConsoleLogger(),
                    new DebugLogger()
                };

            if (args.Length < 3 
                || string.IsNullOrWhiteSpace(args[0])
                || string.IsNullOrWhiteSpace(args[1])
                || string.IsNullOrWhiteSpace(args[2]))
            {
                _loggers?.ForEach(logger => logger.WriteLogMessage("There were missing arguments provided to Main."));
                Console.Error.WriteLine("Usage: OfficeApiMediaExtractionTest <path-to-office-file> <acs-endpoint> <acs-api-key>");
            }
            else
            {
                _docPath = args[0];
                _acsConnectionDetails = new AcsConnectionInfo(args[1], args[2]);

                _worker = new Worker(
                    new OfficeDocManager(
                        new LocalFileHandler(),
                        new DocImageHandler()),
                    new AcsImageAnalyzer(_acsConnectionDetails),
                    _loggers);

                _loggers?.ForEach(logger => logger.WriteLogMessage($"Executing program with document at {_docPath}"));
                await _worker.ExecuteProgramAsync(_docPath);
            }

            Console.WriteLine("Press any key to exit... wherever the any key might be.");
            Console.ReadLine();
        }
    }
}