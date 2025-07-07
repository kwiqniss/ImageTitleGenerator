using OfficeApiMediaExtractionTest.DataTypes;

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface IFileHandler
    {
        public FileInteractionResult Copy(string sourcePath, string destinationPath);

        public bool Exists(string filePath);
    }
}
