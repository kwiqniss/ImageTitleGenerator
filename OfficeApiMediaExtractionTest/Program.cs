// See https://aka.ms/new-console-template for more information
using Azure.AI.Vision.ImageAnalysis;
using OfficeApiMediaExtractionTest.AI.Acs;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using OfficeApiMediaExtractionTest.Loggers;
using OfficeApiMediaExtractionTest.Office;
using System.Diagnostics;

namespace OfficeApiMediaExtractionTest
{
    public class Program
    {
        private const string ACS_ENDPOINT =
            "https://nazaravision.cognitiveservices.azure.com/";

        private const string ACS_API_KEY =
            "3zLPzIAfN7RjptLzX0k3mAUF1QeDRF8mtxfvywT8fSOeNPkBs5IQJQQJ99BGACYeBjFXJ3w3AAAFACOGBFkq"; 

        private static IImageAnalyzer? _imageAnalyzer;
        private static IDocManager? _officeDocManager;
        private static List<ILogger>? _loggers;
        private static IWorker? _worker;

        //var docPath = @"C:\Users\dhollowe\Documents\OfficeDocImagePlayground\sample.docx";
        //var docPath = @"C:\Users\dhollowe\Documents\OfficeDocImagePlayground\sample.pptx";
        //var docPath = @"C:\Users\dhollowe\Documents\OfficeDocImagePlayground\sample.xlsx";
        public static async Task Main(string[] args)
        {
            //TODO: accept endpoint and key as arguments from the command line.
            _imageAnalyzer = new AcsImageAnalyzer(ACS_ENDPOINT, ACS_API_KEY);
            _officeDocManager = new OfficeDocManager();
            _loggers = new List<ILogger>{ new ConsoleLogger(), new DebugLogger() };
            _worker = new Worker(_officeDocManager, _imageAnalyzer, _loggers);

            string docPath;
            if (args.Length < 1 || string.IsNullOrWhiteSpace(args[0]))
            {
                _loggers.ForEach(logger => logger.WriteLogMessage("There was no doc path provided in the arguments given to Main."));
                Console.Error.WriteLine("Usage: OfficeApiMediaExtractionTest <path-to-office-file>");
            }
            else
            {
                docPath = args[0];
                _loggers.ForEach(logger => logger.WriteLogMessage($"Executing program with document at {docPath}"));
                await _worker.ExecuteProgramAsync(docPath);
            }

            // TODO: api endpoint and key should also be added as arguments.
            // TODO: accept multiple file names from the commmand line.
            Console.WriteLine("Press any key to exit... wherever the any key might be.");
            Console.ReadLine();
        }
    }
}