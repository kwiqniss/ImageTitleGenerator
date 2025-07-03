using OfficeApiMediaExtractionTest.AI.Acs;
using OfficeApiMediaExtractionTest.Interfaces;
using OfficeApiMediaExtractionTest.Loggers;
using OfficeApiMediaExtractionTest.Office;
using Moq;

namespace OfficeApiMediaExtractionTest.UnitTests
{
    [TestClass]
    public sealed class Test1
    {
        private const string ACS_ENDPOINT = "https://nazaravision.cognitiveservices.azure.com/"; //https://<your-vision-api-endpoint>.cognitiveservices.azure.com/";
        private const string ACS_API_KEY = "3zLPzIAfN7RjptLzX0k3mAUF1QeDRF8mtxfvywT8fSOeNPkBs5IQJQQJ99BGACYeBjFXJ3w3AAAFACOGBFkq"; // Your Azure Vision API key

        private readonly IWorker _worker;

        public Test1()
        {
            _worker = new Worker(
                new OfficeDocManager(), 
                new AcsImageAnalyzer(ACS_ENDPOINT, ACS_API_KEY), 
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
