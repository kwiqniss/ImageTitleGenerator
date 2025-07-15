using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.IO
{
    public interface IFileHandler
    {
        public FileInteractionResult Copy(string sourcePath, string destinationPath);

        public bool Exists(string filePath);
    }
}
