using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.AI
{
    public interface IImageAnalyzer
    {
        public Task<bool> AddImageDescriptionsAsync(IEnumerable<DocumentImage> images);
    }
}
