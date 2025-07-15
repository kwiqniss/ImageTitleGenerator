using DocumentFormat.OpenXml.Packaging;
using ImageAnalyzer.DataTypes;
using DocumentFormat.OpenXml.Drawing;
using ImageAnalyzer.Office.OfficeWithOpenXML;

namespace ImageAnalyzer.Office.OfficeWithOpenXML.ImageHandlerImplementations
{
    public class XlsxImageHandler: OfficeDocImageHandler
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

                    foreach (var documentImage in GetDocumentImages(drawingsPart, drawingsPart.ImageParts, sheetIndex))
                        yield return documentImage;
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

                    var blips = GetBlips(drawingsPart.WorksheetDrawing);
                    foreach (var blip in blips)
                    {
                        var relId = GetRelId(blip);
                        if (string.IsNullOrEmpty(relId)) continue;

                        var docImage = imageList.FirstOrDefault(img => img.RelId == relId && img.SheetSlideIndex == sheetIndex);
                        if (docImage == null) continue;

                        UpdateTitleProperty(blip, docImage.Title);
                    }
                    drawingsPart.WorksheetDrawing.Save();
                }
                return new FileInteractionResult { IsSuccess = true };
            }
        }

        protected override void UpdateTitleProperty(
            Blip blip,
            string title)
        {
            var nvdProps = blip.Ancestors<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>()
                            .Select(pic => pic.NonVisualPictureProperties?.NonVisualDrawingProperties)
                            .FirstOrDefault(nv => nv != null);
            if (nvdProps != null)
            {
                nvdProps.Description = string.Empty;
                nvdProps.Title = title;
            }
        }
    }
}
