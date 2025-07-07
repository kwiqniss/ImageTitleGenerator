using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Office.ImageHandlerImplementations
{
    public class DocxImageHandler : OfficeDocImageHandler
    {
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
                var drawings = GetDrawings<DocumentFormat.OpenXml.Wordprocessing.Drawing>(mainPart.Document);

                foreach (var drawing in drawings)
                {
                    var blip = GetBlips(drawing).FirstOrDefault();
                    if (blip == null) continue;

                    var relId = GetRelId(blip);
                    if (relId == null) continue;

                    var docImage = imageList.FirstOrDefault(img => img.RelId == relId);
                    if (docImage == null) continue;

                    if (drawing.Descendants<DocProperties>().FirstOrDefault() is { } docProps) 
                        (docProps.Description, docProps.Title) = (string.Empty, docImage.Title);
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
    }
}
