using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeApiMediaExtractionTest.DataTypes
{
    public class AcsConnectionInfo
    {
        public string Endpoint { get; private set; }
        public string ApiKey { get; private set; }
        public AcsConnectionInfo(string endpoint, string apiKey)
        {
            if (string.IsNullOrWhiteSpace(endpoint)) throw new ArgumentNullException(nameof(endpoint));
            if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentNullException(nameof(apiKey));
            Endpoint = endpoint;
            ApiKey = apiKey;
        }
    }
}
