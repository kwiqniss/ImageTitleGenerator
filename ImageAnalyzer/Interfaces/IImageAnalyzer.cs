using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.Interfaces
{
    public interface IImageAnalyzer
    {
        public Task<bool> AddImageDescriptionsAsync(IEnumerable<DocumentImage> images);
    }
}
