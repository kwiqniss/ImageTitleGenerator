using OfficeApiMediaExtractionTest.AI.Acs;
using OfficeApiMediaExtractionTest.Interfaces;
using OfficeApiMediaExtractionTest.IO;
using OfficeApiMediaExtractionTest.Loggers;
using OfficeApiMediaExtractionTest.Office;
using Moq;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Office.ImageHandlerImplementations;

namespace OfficeApiMediaExtractionTest.UnitTests
{
    [TestClass]
    public sealed class E2ETests
    {
        private const string ACS_ENDPOINT = "https://nazaravision.cognitiveservices.azure.com/"; //https://<your-vision-api-endpoint>.cognitiveservices.azure.com/";
        private const string ACS_API_KEY = "3zLPzIAfN7RjptLzX0k3mAUF1QeDRF8mtxfvywT8fSOeNPkBs5IQJQQJ99BGACYeBjFXJ3w3AAAFACOGBFkq"; // Your Azure Vision API key

        private readonly IWorker _worker;

        public E2ETests()
        {
            //var loggers = new List<ILogger> { new ConsoleLogger(), new DebugLogger() };
            _worker = new Worker(
                new OfficeDocManager(
                    new LocalFileHandler(),
                    new List<IImageHandler> 
                    { 
                        new DocxImageHandler(), 
                        new PptxImageHandler(), 
                        new XlsxImageHandler() 
                    }), 
                new AcsImageAnalyzer(new AcsConnectionInfo(ACS_ENDPOINT, ACS_API_KEY)), 
                new List<ILogger> { new Mock<ILogger>().Object });
        }

        [TestMethod]
        public async Task TestDocxAsync()
        {
            await _worker.ExecuteProgramAsync(@"C:\Users\dhollowe\Documents\OfficeDocImagePlayground\sample.docx");
        }

        [TestMethod]
        public async Task TestPptxAsync()
        {
            await _worker.ExecuteProgramAsync(@"C:\Users\dhollowe\Documents\OfficeDocImagePlayground\sample.pptx");
        }

        [TestMethod]
        public async Task TestXlsxAsync()
        {
            await _worker.ExecuteProgramAsync(@"C:\Users\dhollowe\Documents\OfficeDocImagePlayground\sample.xlsx");
        }
    }
}
