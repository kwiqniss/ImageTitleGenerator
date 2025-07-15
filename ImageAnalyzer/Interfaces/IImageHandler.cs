using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.Interfaces
{
    public interface IImageHandler
    {
        public IEnumerable<DocumentImage> GetImages(string docPath);
        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);
    }
}
