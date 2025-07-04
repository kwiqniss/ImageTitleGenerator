using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Office
{
    public abstract class OfficeDocumentImageHandler: IImageHandler
    {
        public virtual IEnumerable<string> SupportedFileExtensions { get; set; } = [];

        public abstract IEnumerable<DocumentImage> GetImages(string docPath);

        public abstract FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);

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
