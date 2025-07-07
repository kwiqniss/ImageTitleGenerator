using DocumentFormat.OpenXml.Packaging;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Office
{
    public abstract class OfficeDocImageHandler: IImageHandler
    {
        public virtual IEnumerable<string> SupportedFileExtensions { get; set; } = [];

        public abstract IEnumerable<DocumentImage> GetImages(string docPath);

        public abstract FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);

        public string GetIdOfPart(OpenXmlPart xmlPart, ImagePart imagePart)
        {
            return xmlPart.GetIdOfPart(imagePart);
        }

        public IEnumerable<DocumentImage> GetDocumentImages(
            OpenXmlPart xmlPart, 
            IEnumerable<ImagePart> imageParts, 
            int pageIndex = -1)
        {
            int imageIndex = 0;
            foreach (var imagePart in imageParts)
            {
                string extension = GetImageExtension(imagePart.ContentType);
                using var imageStream = imagePart.GetStream();
                var ms = new MemoryStream();
                imageStream.CopyTo(ms);
                ms.Position = 0;

                var relId = GetIdOfPart(xmlPart, imagePart);
                yield return new DocumentImage(
                    extension,
                    ms,
                    relId,
                    $"Sheet{pageIndex}_Image{++imageIndex}",
                    pageIndex
                );
            }
        }

        internal static string GetImageExtension(string contentType)
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
    }
}
