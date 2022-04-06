using System.IO;

namespace WebApplication1.Entities.Dtos
{
    public class DocumentTypeFile
    {
        public MemoryStream File { get; set; }
        public string NomFile { get; set; }
        public string Extension { get; set; }
        public string TypeFile { get; set; }
    }
}
