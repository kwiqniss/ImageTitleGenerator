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
    }
}
