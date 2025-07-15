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
        private readonly IWorker _worker;

        public E2ETests()
        {
            // Build configuration to include user secrets
            var configuration = new ConfigurationBuilder()
                .AddUserSecrets<E2ETests>() // Uses the UserSecretsId from your .csproj
                .Build();

            _worker = new Worker(
                new DocManager(
                    new LocalFileHandler(),
                    new DocImageHandler()),
                new AcsImageAnalyzer(new AcsConnectionInfo(ACS_ENDPOINT, configuration["Acs:ApiKey"])),
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
