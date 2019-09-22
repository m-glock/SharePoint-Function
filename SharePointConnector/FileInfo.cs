using System.Collections.Generic;

namespace SharePointConnector
{
    class FileDto
    {
        public string collectionName { get; set; }
        public Dictionary<string, Dictionary<string, string>> Files { get; set; }

    }
}