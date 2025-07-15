using Microsoft.Extensions.Configuration;
using Moq;
using ImageAnalyzer.AI.Acs;
using ImageAnalyzer.IO;
using ImageAnalyzer.DocumentInteractions;
using ImageAnalyzer.DataTypes;
using ImageAnalyzerProgram.Loggers;

namespace ImageAnalyzerProgram.Tests
{
    [TestClass]
    public sealed class E2ETests
    {
        private const string ACS_ENDPOINT = "https://nazaravision.cognitiveservices.azure.com/";
        private readonly IConfiguration _configuration;
        private readonly IWorker _worker;

        public E2ETests()
        {
            // Build configuration to include user secrets
            _configuration = new ConfigurationBuilder()
                .AddUserSecrets<E2ETests>() // Uses the UserSecretsId from your .csproj
                .Build();

            // Example: Access a secret value
            string apiKey = _configuration["Acs:ApiKey"];

            _worker = new Worker(
                new DocManager(
                    new LocalFileHandler(),
                    new DocImageHandler()),
                new AcsImageAnalyzer(new AcsConnectionInfo(ACS_ENDPOINT, apiKey)),
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
