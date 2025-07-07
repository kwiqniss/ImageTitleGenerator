using ImageAnalyzer.DataTypes;
using ImageAnalyzer.Interfaces;

namespace ImageAnalyzer.IO
{
    public class LocalFileHandler: IFileHandler
    {
        public bool Exists(string filePath)
        {
            return File.Exists(filePath);
        }

        public FileInteractionResult Copy(string sourcePath, string destinationPath)
        {
            try
            {
                File.Copy(sourcePath, destinationPath, true);
                return new FileInteractionResult
                {
                    IsSuccess = true,
                    Value = destinationPath
                };
            }
            catch (Exception ex)
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Value = ex.Message,
                    InteractionException = ex
                };
            }
        }
    }
}
