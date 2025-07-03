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
            return extension switch
            {
                ".docx" => GetImagesFromDocx(docPath),
                ".pptx" => GetImagesFromPptx(docPath),
                ".xlsx" => GetImagesFromXlsx(docPath),
                _ => Enumerable.Empty<DocumentImage>()
            };
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
            try
            {
                return extension switch
                {
                    ".docx" => SaveDocxImageTitles(images, docPath),
                    ".pptx" => SavePptxImageTitles(images, docPath),
                    ".xlsx" => SaveXlsxImageTitles(images, docPath),
                    _ => new FileInteractionResult
                    {
                        IsSuccess = false,
                        Message = $"Unsupported file type: {extension}"
                    }
                };
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

        private static IEnumerable<DocumentImage> GetImagesFromDocx(string docxPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, false))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null)
                    yield break;

                int imageIndex = 0;
                foreach (var imagePart in mainPart.ImageParts)
                {
                    string extension = GetImageExtension(imagePart.ContentType);
                    using var imageStream = imagePart.GetStream();
                    var ms = new MemoryStream();
                    imageStream.CopyTo(ms);
                    ms.Position = 0;

                    var relId = mainPart.GetIdOfPart(imagePart);
                    yield return new DocumentImage(
                        extension,
                        ms,
                        relId,
                        $"Image {++imageIndex}"
                    );
                }
            }
        }

        private static IEnumerable<DocumentImage> GetImagesFromPptx(string pptxPath)
        {
            using (PresentationDocument pptDoc = PresentationDocument.Open(pptxPath, false))
            {
                var presentationPart = pptDoc.PresentationPart;
                if (presentationPart == null)
                    yield break;

                int slideIndex = 0;
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    slideIndex++;
                    int imageIndex = 0;
                    foreach (var imagePart in slidePart.ImageParts)
                    {
                        string extension = GetImageExtension(imagePart.ContentType);
                        using var imageStream = imagePart.GetStream();
                        var ms = new MemoryStream();
                        imageStream.CopyTo(ms);
                        ms.Position = 0;

                        var relId = slidePart.GetIdOfPart(imagePart);
                        yield return new DocumentImage(
                            extension,
                            ms,
                            relId,
                            $"Slide{slideIndex}_Image{++imageIndex}",
                            slideIndex
                        );
                    }
                }
            }
        }

        private static IEnumerable<DocumentImage> GetImagesFromXlsx(string xlsxPath)
        {
            using (SpreadsheetDocument xlsDoc = SpreadsheetDocument.Open(xlsxPath, false))
            {
                var workbookPart = xlsDoc.WorkbookPart;
                if (workbookPart == null)
                    yield break;

                int sheetIndex = 0;
                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    sheetIndex++;
                    var drawingsPart = worksheetPart.DrawingsPart;
                    if (drawingsPart == null)
                        continue;

                    int imageIndex = 0;
                    foreach (var imagePart in drawingsPart.ImageParts)
                    {
                        string extension = GetImageExtension(imagePart.ContentType);
                        using var imageStream = imagePart.GetStream();
                        var ms = new MemoryStream();
                        imageStream.CopyTo(ms);
                        ms.Position = 0;

                        var relId = drawingsPart.GetIdOfPart(imagePart);
                        yield return new DocumentImage(
                            extension,
                            ms,
                            relId,
                            $"Sheet{sheetIndex}_Image{++imageIndex}",
                            sheetIndex
                        );
                    }
                }
            }
        }

        private static string GetImageExtension(string contentType)
        {
            return contentType switch
            {
                "image/png" => ".png",
                "image/jpeg" => ".jpg",
                "image/gif" => ".gif",
                "image/bmp" => ".bmp",
                "image/tiff" => ".tiff",
                _ => ".img"
            };
        }

        private static FileInteractionResult SaveDocxImageTitles(IEnumerable<DocumentImage> images, string docxPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docxPath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null)
                {
                    return new FileInteractionResult
                    {
                        IsSuccess = false,
                        Message = "doc MainDocumentPart was null."
                    };
                }

                var imageList = images.ToList();
                var drawings = mainPart.Document.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>();

                foreach (var drawing in drawings)
                {
                    var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                    if (blip?.Embed == null) continue;
                    var relId = blip.Embed.Value;
                    var docImage = imageList.FirstOrDefault(img => img.RelId == relId);
                    if (docImage == null) continue;

                    var docProps = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>().FirstOrDefault();
                    if (docProps == null)
                    {
                        docProps = new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties();
                        drawing.Append(docProps);
                    }
                    docProps.Description = string.Empty; 
                    docProps.Title = docImage.Title;
                }

                try
                {
                    mainPart.Document.Save();
                    return new FileInteractionResult { IsSuccess = true };
                }
                catch (Exception ex)
                {
                    return new FileInteractionResult
                    {
                        IsSuccess = false,
                        Message = $"Failed to save document: {ex.Message}",
                        InteractionException = ex
                    };
                }
            }
        }

        private static FileInteractionResult SavePptxImageTitles(IEnumerable<DocumentImage> images, string pptxPath)
        {
            using (PresentationDocument pptDoc = PresentationDocument.Open(pptxPath, true))
            {
                var imageList = images.ToList();
                var presentationPart = pptDoc.PresentationPart;
                if (presentationPart == null)
                {
                    return new FileInteractionResult
                    {
                        IsSuccess = false,
                        Message = "pptx PresentationPart was null."
                    };
                }

                int slideIndex = 0;
                foreach (var slidePart in presentationPart.SlideParts)
                {
                    slideIndex++;
                    var blips = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Blip>();
                    foreach (var blip in blips)
                    {
                        var relId = blip.Embed?.Value;
                        if (string.IsNullOrEmpty(relId)) continue;
                        var docImage = imageList.FirstOrDefault(img => img.RelId == relId && img.SheetSlideIndex == slideIndex);
                        if (docImage == null) continue;

                        // Try Drawing.Pictures.Picture first
                        var picNvProps = blip.Ancestors<DocumentFormat.OpenXml.Drawing.Pictures.Picture>()
                            .Select(pic => pic.NonVisualPictureProperties?.NonVisualDrawingProperties)
                            .FirstOrDefault(nv => nv != null);
                        if (picNvProps != null)
                        {
                            picNvProps.Description = string.Empty;
                            picNvProps.Title = docImage.Title;
                            continue;
                        }

                        // Try Presentation.Picture second
                        var presNvProps = blip.Ancestors<DocumentFormat.OpenXml.Presentation.Picture>()
                            .Select(pic => pic.NonVisualPictureProperties?.NonVisualDrawingProperties)
                            .FirstOrDefault(nv => nv != null);
                        if (presNvProps != null)
                        {
                            presNvProps.Description = string.Empty;
                            presNvProps.Title = docImage.Title;
                        }
                    }
                    slidePart.Slide.Save();
                }
                return new FileInteractionResult { IsSuccess = true };
            }
        }

        private static FileInteractionResult SaveXlsxImageTitles(IEnumerable<DocumentImage> images, string xlsxPath)
        {
            using (SpreadsheetDocument xlsDoc = SpreadsheetDocument.Open(xlsxPath, true))
            {
                var imageList = images.ToList();
                var workbookPart = xlsDoc.WorkbookPart;
                if (workbookPart == null)
                {
                    return new FileInteractionResult
                    {
                        IsSuccess = false,
                        Message = "xlsx WorkbookPart was null."
                    };
                }

                int sheetIndex = 0;
                foreach (var worksheetPart in workbookPart.WorksheetParts)
                {
                    sheetIndex++;
                    var drawingsPart = worksheetPart.DrawingsPart;
                    if (drawingsPart == null)
                        continue;

                    var pictures = drawingsPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>();
                    foreach (var pic in pictures)
                    {
                        var blip = pic.BlipFill?.Blip;
                        var relId = blip?.Embed?.Value;
                        if (string.IsNullOrEmpty(relId)) continue;
                        var docImage = imageList.FirstOrDefault(img => img.RelId == relId && img.SheetSlideIndex == sheetIndex);
                        if (docImage == null) continue;

                        var nvdProps = pic.NonVisualPictureProperties?.NonVisualDrawingProperties;
                        if (nvdProps == null)
                        {
                            nvdProps = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties();
                            pic.Append(nvdProps);
                        }
                        nvdProps.Description = string.Empty;
                        nvdProps.Title = docImage.Title;
                    }
                    drawingsPart.WorksheetDrawing.Save();
                }
                return new FileInteractionResult { IsSuccess = true };
            }
        }
    }
}