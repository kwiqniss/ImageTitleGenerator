using OfficeApiMediaExtractionTest.DataTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface IImageHandler
    {
        public IEnumerable<string> SupportedFileExtensions { get; }
        public IEnumerable<DocumentImage> GetImages(string docPath);
        public FileInteractionResult SaveImageTitles(IEnumerable<DocumentImage> images, string docPath);
    }
}
