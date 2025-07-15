using Microsoft.Extensions.Configuration;
using Moq;
using ImageAnalyzer.AI.Acs;
using ImageAnalyzer.IO;
using ImageAnalyzer.DocumentInteractions;
using ImageAnalyzer.DataTypes;
using ImageAnalyzerProgram.Loggers;
using System.Diagnostics;

namespace ImageAnalyzerProgram.Tests
{
    // TODO: Must setup user secrets.json file and include the following keys: "Acs:Endpoint", "Acs:ApiKey" -https://learn.microsoft.com/en-us/aspnet/core/security/app-secrets

    [TestClass]
    public sealed class E2ETests
    {
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
                new AcsImageAnalyzer(new AcsConnectionInfo(configuration["Acs:Endpoint"], configuration["Acs:ApiKey"])),
                new List<ILogger> { new Mock<ILogger>().Object });

            Process.Start("explorer.exe", "SampleFiles"); // Open the SampleFiles directory in File Explorer for easy access
        }

        [TestMethod]
        public async Task TestDocxAsync()
        {
            await _worker.ExecuteProgramAsync(@"SampleFiles\sample.docx");
        }

        [TestMethod]
        public async Task TestPptxAsync()
        {
            await _worker.ExecuteProgramAsync(@"SampleFiles\sample.pptx");
        }

        [TestMethod]
        public async Task TestXlsxAsync()
        {
            await _worker.ExecuteProgramAsync(@"SampleFiles\sample.xlsx");
        }
    }
}
