using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.Interfaces
{
    public interface IDocManager
    {
        public FileInteractionResult CopyDoc(string sourcePath);
        public IEnumerable<DocumentImage> GetImages(string docPath);
        public FileInteractionResult SaveTitles(IEnumerable<DocumentImage> images, string docPath);
    }
}
