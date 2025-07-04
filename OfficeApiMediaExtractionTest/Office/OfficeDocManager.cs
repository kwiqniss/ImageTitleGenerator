using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using Path = System.IO.Path;

namespace OfficeApiMediaExtractionTest.Office
{
    /* TODO: [dhollowe] we can improve the performance by making sure we don't query every image in the collection 
     * to compare relId and sheetslideId but instead use them as indexes so we can just loop through the collection once when saving titles.
     */
    public class OfficeDocManager : IDocManager
    {
        private readonly IEnumerable<IImageHandler> _imageHandlers;
        public OfficeDocManager(IEnumerable<IImageHandler> imageHandlers)
        {
            _imageHandlers = imageHandlers;
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

            if (!File.Exists(sourcePath))
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
            } while (File.Exists(destPath));

            try
            {
                File.Copy(sourcePath, destPath);
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
            var extension = Path.GetExtension(docPath).ToLowerInvariant();
            var imageHandler = _imageHandlers.FirstOrDefault(handler => handler.SupportedFileExtensions.Contains(extension));
            return imageHandler == null 
                ? Enumerable.Empty<DocumentImage>() 
                : imageHandler.GetImages(docPath);
        }

        public FileInteractionResult SaveTitles(IEnumerable<DocumentImage> images, string docPath)
        {
            if (images == null || string.IsNullOrWhiteSpace(docPath))
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Message = "No images in doc."
                };
            }

            var extension = Path.GetExtension(docPath).ToLowerInvariant();
            var imageHandler = _imageHandlers.FirstOrDefault(handler =>
                handler.SupportedFileExtensions.Contains(extension));

            try
            {
                return imageHandler == null 
                    ? new FileInteractionResult
                        {
                            IsSuccess = false,
                            Message = $"No handler found for file type: {extension}"
                        }
                    : imageHandler.SaveImageTitles(images, docPath);
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