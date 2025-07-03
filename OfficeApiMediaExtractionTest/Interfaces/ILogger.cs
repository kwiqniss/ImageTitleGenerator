using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.Interfaces
{
    public interface ILogger
    {
        void WriteLogMessage(string message);
        void WriteLogError(string? message, Exception? ex = null);
    }
}
