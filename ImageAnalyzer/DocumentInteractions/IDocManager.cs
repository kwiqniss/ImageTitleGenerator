using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.DocumentInteractions
{
    public interface IDocManager
    {
        public FileInteractionResult CopyDoc(string sourcePath);
        public IEnumerable<DocumentImage> GetImages(string docPath);
        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);
    }
}
