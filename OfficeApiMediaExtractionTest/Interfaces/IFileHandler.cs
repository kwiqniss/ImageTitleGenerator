using OfficeApiMediaExtractionTest.DataTypes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface IFileHandler
    {
        public FileInteractionResult Copy(string sourcePath, string destinationPath);

        public bool Exists(string filePath);
    }
}
