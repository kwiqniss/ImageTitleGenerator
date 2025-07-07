using OfficeApiMediaExtractionTest.DataTypes;

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface IImageHandler
    {
        public IEnumerable<string> SupportedFileExtensions { get; }
        public IEnumerable<DocumentImage> GetImages(string docPath);
        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);
    }
}
