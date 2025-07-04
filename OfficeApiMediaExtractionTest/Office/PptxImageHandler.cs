using DocumentFormat.OpenXml.Packaging;
using OfficeApiMediaExtractionTest.DataTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Office
{
    public class PptxImageHandler: OfficeDocImageHandler
    {
        public override IEnumerable<string> SupportedFileExtensions { get; set; } = [".pptx"];

        public override IEnumerable<DocumentImage> GetImages(string docPath)
        {
            using (PresentationDocument pptDoc = PresentationDocument.Open(docPath, false))
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

        public override FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath)
        {
            using (PresentationDocument pptDoc = PresentationDocument.Open(docPath, true))
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
    }
}
