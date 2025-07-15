using ImageAnalyzer.DataTypes;
using ImageAnalyzer.IO;
using Path = System.IO.Path;

namespace ImageAnalyzer.DocumentInteractions
{
    public class DocManager : IDocManager
    {
        private readonly IFileHandler _fileHandler;
        private readonly IImageHandler _imageHandler;
        
        public DocManager(IFileHandler fileHandler, IImageHandler imageHandler)
        {
            _fileHandler = fileHandler;
            _imageHandler = imageHandler;
        }
        
        public FileInteractionResult CopyDoc(string sourcePath)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Message = "Source path is null or empty/whitespace."
                };
            }

            if (!_fileHandler.Exists(sourcePath))
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Message = "Source file does not exist."
                };
            }

            string directory = Path.GetDirectoryName(sourcePath) ?? "";
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(sourcePath);
            string extension = Path.GetExtension(sourcePath);

            int copyIndex = 1;
            string destPath;
            do
            {
                destPath = Path.Combine(directory, $"{fileNameWithoutExt} (Copy {copyIndex}){extension}");
                copyIndex++;
            } while (_fileHandler.Exists(destPath));

            try
            {
                _fileHandler.Copy(sourcePath, destPath);
                return new FileInteractionResult
                {
                    IsSuccess = true,
                    Value = destPath
                };
            }
            catch (Exception ex)
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Message = $"Copy failed: {ex.Message}",
                    InteractionException = ex
                };
            }
        }

        public IEnumerable<DocumentImage> GetImages(string docPath)
        {
            return _imageHandler.GetImages(docPath);
        }

        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath)
        {
            if (images == null || string.IsNullOrWhiteSpace(docPath))
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Message = "No images in doc."
                };
            }

            try
            {
                return _imageHandler.SaveImageTitles(images, docPath);
            }
            catch (Exception ex)
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Message = $"Exception: {ex.Message}",
                    InteractionException = ex
                };
            }
        }
    }
}