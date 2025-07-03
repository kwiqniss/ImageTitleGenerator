using OfficeApiMediaExtractionTest.DataTypes;

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface IImageAnalyzer
    {
        public Task<bool> AddImageDescriptionsAsync(IEnumerable<DocumentImage> images);
    }
}
