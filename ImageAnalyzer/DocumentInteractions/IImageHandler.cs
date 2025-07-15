using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.DocumentInteractions
{
    public interface IImageHandler
    {
        public IEnumerable<DocumentImage> GetImages(string docPath);
        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);
    }
}
