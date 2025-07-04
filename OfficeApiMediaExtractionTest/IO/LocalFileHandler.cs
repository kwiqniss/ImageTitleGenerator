using DocumentFormat.OpenXml.Vml;
using OfficeApiMediaExtractionTest.DataTypes;
using OfficeApiMediaExtractionTest.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.IO
{
    public class LocalFileHandler: IFileHandler
    {
        public bool Exists(string filePath)
        {

            return File.Exists(filePath);
        }

        public FileInteractionResult Copy(string sourcePath, string destinationPath)
        {
            try
            {
                File.Copy(sourcePath, destinationPath, true);
                return new FileInteractionResult
                {
                    IsSuccess = true,
                    Value = destinationPath
                };
            }
            catch (Exception ex)
            {
                return new FileInteractionResult
                {
                    IsSuccess = false,
                    Value = ex.Message,
                    InteractionException = ex
                };
            }
        }
    }
}
