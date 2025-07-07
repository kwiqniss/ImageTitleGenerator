using Azure;
using Azure.AI.Vision.ImageAnalysis;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using System.Net;

namespace OfficeApiMediaExtractionTest.AI.Acs
{
    public class AcsImageAnalyzer : IImageAnalyzer
    {
        private readonly ImageAnalysisClient _client;

        public AcsImageAnalyzer(AcsConnectionInfo connectionDetails)
        {
            var endpoint = connectionDetails.Endpoint;
            var apiKey = connectionDetails.ApiKey;

            if (string.IsNullOrWhiteSpace(endpoint)) throw new ArgumentNullException(nameof(endpoint));
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentNullException(nameof(apiKey));

            _client = new ImageAnalysisClient(new Uri(endpoint), new AzureKeyCredential(apiKey));
        }

        public async Task<bool> AddImageDescriptionsAsync(IEnumerable<DocumentImage> images)
        {
            if (_client == null || images == null || images.Count() < 1)
                return false;

            foreach (var image in images)
            {
                if (image != null && image.ImageStream != null)
                    image.Title = await DescribeImageAsync(image.ImageStream);
            }

            return true; // to indicate success
        }

        private async Task<string> DescribeImageAsync(Stream imageStream)
        {
            var imageBinaryData = BinaryData.FromStream(imageStream);
            var result = await _client.AnalyzeAsync(
                imageBinaryData,
                VisualFeatures.Caption);
            
            return result?.Value?.Caption?.Text ?? "No description found.";
        }
    }
}
