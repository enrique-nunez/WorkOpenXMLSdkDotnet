using Microsoft.AspNetCore.Mvc;
using System;
using System.IO;
using System.Threading.Tasks;
using WebApplication1.Entities.Constants;
using WebApplication1.Entities.Dtos;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumentFormatController : Controller
    {

        [HttpPost("export")]
        public async Task<TablaReports> ExportDocumentDoc()
        {
            WordDocumentService _wordDocumentService = new WordDocumentService();
            TablaReports tablaReport = _wordDocumentService.GetTableReportsByType();
            var pathDocument = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Template", "Formato Construcción OCT.docx");
            var pathDocumentGenerate = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "Upload", $"Export_{Guid.NewGuid()}.docx");
            _wordDocumentService.DuplicateDocument(pathDocument, pathDocumentGenerate);
            _wordDocumentService.DocumentStampingByReportType(pathDocumentGenerate, tablaReport);
            
            return tablaReport;
        }
    }
}
