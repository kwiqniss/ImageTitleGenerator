using DocumentFormat.OpenXml.Packaging;
using OfficeApiMediaExtractionTest.DataTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Office
{
    public class XlsxImageHandler: OfficeDocumentImageHandler
    {
        public override IEnumerable<string> SupportedFileExtensions { get; set; } = [".xlsx"];

        public override IEnumerable<DocumentImage> GetImages(string docPath)
        {
            using (SpreadsheetDocument xlsDoc = SpreadsheetDocument.Open(docPath, false))
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

        public override FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath)
        {
            using (SpreadsheetDocument xlsDoc = SpreadsheetDocument.Open(docPath, true))
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
