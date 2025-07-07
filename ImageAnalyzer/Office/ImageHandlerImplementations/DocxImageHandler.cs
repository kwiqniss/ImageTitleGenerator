using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ImageAnalyzer.DataTypes;

namespace ImageAnalyzer.Office.ImageHandlerImplementations
{
    public class DocxImageHandler : OfficeDocImageHandler
    {
        private OpenXmlCompositeElement? _activeDrawing { get; set; }

        public override IEnumerable<string> SupportedFileExtensions { get; set; } = [".docx"];

        public override IEnumerable<DocumentImage> GetImages(string docPath)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(docPath, false))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null)
                    yield break;

                foreach (var documentImage in GetDocumentImages(mainPart, mainPart.ImageParts))
                    yield return documentImage;
            }
        }

        public override FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docxPath)
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
                var drawings = GetDrawings<Drawing>(mainPart.Document);

                foreach (var drawing in drawings)
                {
                    if (drawing == null) continue;

                    var blip = GetBlips(drawing).FirstOrDefault();
                    if (blip == null) continue;

                    var relId = GetRelId(blip);
                    if (relId == null) continue;

                    var docImage = imageList.FirstOrDefault(img => img.RelId == relId);
                    if (docImage == null) continue;

                    UpdateTitleProperty(drawing, blip, docImage.Title);
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

        protected override void UpdateTitleProperty(Blip blip, string title)
        {
            if (_activeDrawing?.Descendants<DocProperties>().FirstOrDefault() is { } docProps)
                (docProps.Description, docProps.Title) = (string.Empty, title);
        }

        private void UpdateTitleProperty(OpenXmlCompositeElement drawing, Blip blip, string title)
        {
            _activeDrawing = drawing;
            UpdateTitleProperty(blip, title);
            _activeDrawing = null;
        }
    }
}
