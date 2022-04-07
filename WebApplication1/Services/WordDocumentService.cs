using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using WebApplication1.Entities.Constants;
using WebApplication1.Entities.Dtos;
using WebApplication1.Entities.DynamonDB;
using WebApplication1.Helpers;

namespace WebApplication1.Services
{
    public class WordDocumentService
    {
        private readonly WordDocumentHelper _wordDocumentHelper;
        public WordDocumentService()
        {
            _wordDocumentHelper = new WordDocumentHelper();
        }
        public WordDocumentService(WordDocumentHelper wordDocumentHelper)
        {
            _wordDocumentHelper = wordDocumentHelper;
        }

        public TablaReports GetTableReportsByType()
        {
            TablaReports tablaReports = new TablaReports();

            tablaReports.risk = new Risk()
            {
                NumTramite = "01111",
                NumRiesgo= "706baee4-facb-42ec-af9d-132f4ec0f0df",
                DesExposic= "exposicion1",
                DesGeologico= "se han cambiado los datos de GEO",
                DesGeotecnic= "actaulizacion de info",
                DesHuelgaCc= null,
                DesIncendio= "se quemo",
                DesInundacion= "inundacion1",
                DesLluvia= "lluvia1",
                DesMaremoto= "maremoto1",
                DesPerdBen= "camion",
                DesRespCivil= "respcivil1",
                DesRobo= null,
                DesRotMaq= null,
                DesTerremoto= "terremoto1",
                IndDel= "0"
            };

            tablaReports.sinister = new Sinister()
            {
                NumTramite =  "01111",
                NumSiniestro =  "957854",
                DesSiniestro =  "Descripcion de siniestro",
                DesTieParal =  "Descripcion de Tie Parala",
                FecSiniestro =  "12/01/2022",
                IndDel =  null
            };

            tablaReports.inspection = new Inspection()
            {
                CorreoInsp = "mauvigil@gmail.com",
                NumTramite = "01111",
                ActividadLocal = null,
                CodCordMatriz = null,
                CodEstado = null,
                CodEstatus = null,
                CodEstructura = null,
                CodInspector = null,
                CodTipRiesgo = null,
                CodTipoInforme = "01",
                CodUso = null,
                CodAgente = null,
                CodEstadoTramite = null,
                CodTipoCotizacion = null,
                CodTipoDocumento = null,
                CodUsuario = null,
                Colindantes = null,
                DesDireccion = "Av 28 de Julio, Miraflores",
                DesEdifInst = null,
                DesEstructura = null,
                DesGarantia = null,
                DesGeneral = null,
                DesHuelgaCc = null,
                DesIncendio = null,
                DesJustificacion = null,
                DesMant = null,
                DesNueInver = null,
                DesObservacion = null,
                DesPerdBen = null,
                DesRazonsocial = "Stefanini Group Solutions",
                DesRobo = null,
                DesRotMaq = null,
                DesSumario = null,
                Descripcion = null,
                FecActualizacion = "2022-01-24",
                FecInspeccion = "14/02/2022",
                FecModif = null,
                FecSolicitud = null,
                FecVenInsp = null,
                FechaActualizacion = null,
                FechaCarga = null,
                FechaEfecto = null,
                FechaInspeccion = null,
                FechaSolicitud = null,
                FechaVencimiento = null,
                HorInspeccion = null,
                IdTarea = null,
                MtoContenido = null,
                MtoEdificacion = null,
                MtoExistencia = null,
                MtoLucro = null,
                MtoMaquinaria = null,
                Mto_total = null,
                NomInspector = null,
                NomRazsocial = null,
                NombreAgente = null,
                NombreEstadoTramite = null,
                NombreUsuarioInspector = "Enrique Nuñez",
                NumInforme = null,
                NumLatitud = null,
                NumLongitud = null,
                NumPiso = null,
                NumSotano = null,
                NumeroDocumento = null,
                Oficina = null,
                PorcenEml = null,
                PorcenPml = null,
                RiosQuebradas = null,
                TipoCotizacion = null,
                TipoDocumento = null
            };

            tablaReports.interviewed = new Interviewed()
            {
                NumTramite = "01111",
                NumEntrevistado = "6a132375-7b12-4dcb-9e99-7411e9ced5248",
                DesCargo = "Administrador",
                DesCorreo = "admin@gmail.com",
                DesTelefono = "957456852",
                IndDel = "15",
                NomEntrevistado = "Abel Miraval"
            };

            Archive archive1 = new Archive()
            {
                NumTramite = "01111",
                NumArchivo = "91b15439-05c2-4d24-b09b-a8a359786469",
                NomArchivo = "logo_stefanini",
                CntPesoArchivo = "6751",
                DesMineType = "png",
                FecCrea = "2022-03-30T16:40:07.199Z",
                IndDel = null,
                TipArchivo = "01"
            };

            Archive archive2 = new Archive()
            {
                NumTramite = "01111",
                NumArchivo = "91b15439-05c2-4d24-b09b-a8a359786469",
                NomArchivo = "logo_stefanini",
                CntPesoArchivo = "6751",
                DesMineType = "png",
                FecCrea = "2022-03-30T16:40:07.199Z",
                IndDel = null,
                TipArchivo = "01"
            };

            tablaReports.listArchive = new List<Archive>();
            tablaReports.listArchive.Add(archive1);
            tablaReports.listArchive.Add(archive2);
            Console.WriteLine(tablaReports);
            return tablaReports;
        }

        public bool DocumentStampingByReportType(string pathDocumentGenerate, TablaReports tablaReports)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(pathDocumentGenerate, true))
            {
                _wordDocumentHelper.WritingWordReportConstruccionInspection(doc, tablaReports.inspection);
                _wordDocumentHelper.WritingWordReportConstruccionInterviewed(doc, tablaReports.interviewed);
                _wordDocumentHelper.WritingWordReportConstruccionRisk(doc, tablaReports.risk);
                _wordDocumentHelper.WritingWordReportConstruccionSinister(doc, tablaReports.sinister);
                List<DocumentTypeFile> listDocumentTypeFiles = ListFilesToS3(tablaReports);
                _wordDocumentHelper.AddImgToWordReportConstruccionArchive(doc, listDocumentTypeFiles);
                _wordDocumentHelper.AddTableWordReportConstruccionArchive(doc, listDocumentTypeFiles);
                _wordDocumentHelper.AddImgToTableToWordReportConstruccion(doc, listDocumentTypeFiles);
                doc.Close();
            }

            return true;
        }

        public void DuplicateDocument(string pathDocument, string pathDocumentGenerate)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(pathDocument, true)) //open source word file
            {
                Document document = doc.MainDocumentPart.Document;
                OpenXmlPackage res = doc.SaveAs(pathDocumentGenerate); // copy it to outfile path for editing
                res.Close();
            }
        }

        public MemoryStream GetFileToBase64(string base64String)
        {
            byte[] bytes = Convert.FromBase64String(base64String);
            MemoryStream ms = new MemoryStream(bytes);
            return ms;
        }

        public List<DocumentTypeFile> ListFilesToS3(TablaReports tablaReports)
        {
            try
            {
                MemoryStream msImageLogo = GetFileToBase64(ImageBase64.IMAGE_LOGO);
                MemoryStream msImageFirma = GetFileToBase64(ImageBase64.IMAGE_FIRMA);
                MemoryStream msImageFotografia_1 = GetFileToBase64(ImageBase64.IMAGE_FOTOGRAFIA_1);
                MemoryStream msImageFotografia_2 = GetFileToBase64(ImageBase64.IMAGE_FOTOGRAFIA_2);
                MemoryStream msImageFotografia_3 = GetFileToBase64(ImageBase64.IMAGE_FOTOGRAFIA_3);
                DocumentTypeFile document1 = new DocumentTypeFile()
                {
                    File = msImageLogo,
                    Extension = "png",
                    NomFile = "logo_stefanini.png",
                    TypeFile = "01"
                };
                DocumentTypeFile document2 = new DocumentTypeFile()
                {
                    File = msImageFirma,
                    Extension = "png",
                    NomFile = "firma.png",
                    TypeFile = "02"
                };
                DocumentTypeFile document3 = new DocumentTypeFile()
                {
                    File = msImageFotografia_1,
                    Extension = "jpg",
                    NomFile = "fotografia_1.jpg",
                    TypeFile = "04"
                };
                DocumentTypeFile document4 = new DocumentTypeFile()
                {
                    File = msImageFotografia_2,
                    Extension = "jpg",
                    NomFile = "fotografia_2.jpg",
                    TypeFile = "04"
                };
                DocumentTypeFile document5 = new DocumentTypeFile()
                {
                    File = msImageFotografia_3,
                    Extension = "jpg",
                    NomFile = "fotografia_3.jpg",
                    TypeFile = "04"
                };
                List<DocumentTypeFile> listDocumentTypeFiles = new List<DocumentTypeFile>();
                listDocumentTypeFiles.Add(document1);
                listDocumentTypeFiles.Add(document2);
                listDocumentTypeFiles.Add(document3);
                listDocumentTypeFiles.Add(document4);
                listDocumentTypeFiles.Add(document5);
                return listDocumentTypeFiles;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }

        }
    
    
    }
}
