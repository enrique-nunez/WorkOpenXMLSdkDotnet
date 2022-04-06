using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using WebApplication1.Entities.Constants;
using WebApplication1.Entities.Dtos;
using WebApplication1.Entities.DynamonDB;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace WebApplication1.Helpers
{
    public class WordDocumentHelper
    {
        private class WordMatchedPhrase
        {
            public int charStartInFirstPar { get; set; }
            public int charEndInLastPar { get; set; }

            public int firstCharParOccurance { get; set; }
            public int lastCharParOccurance { get; set; }
        }

        #region REPLACE STRING
        public WordprocessingDocument ReplaceStringInWordDocumennt(WordprocessingDocument wordprocessingDocument, string replaceWhat, string replaceFor)
        {
            List<WordMatchedPhrase> matchedPhrases = FindWordMatchedPhrases(wordprocessingDocument, replaceWhat);

            Document document = wordprocessingDocument.MainDocumentPart.Document;
            int i = 0;
            bool isInPhrase = false;
            bool isInEndOfPhrase = false;
            foreach (Text text in document.Descendants<Text>()) // <<< Here
            {
                char[] textChars = text.Text.ToCharArray();
                List<WordMatchedPhrase> curParPhrases = matchedPhrases.FindAll(a => (a.firstCharParOccurance.Equals(i) || a.lastCharParOccurance.Equals(i)));
                StringBuilder outStringBuilder = new StringBuilder();

                for (int c = 0; c < textChars.Length; c++)
                {
                    if (isInEndOfPhrase)
                    {
                        isInPhrase = false;
                        isInEndOfPhrase = false;
                    }

                    foreach (var parPhrase in curParPhrases)
                    {
                        if (c == parPhrase.charStartInFirstPar && i == parPhrase.firstCharParOccurance)
                        {
                            outStringBuilder.Append(replaceFor);
                            isInPhrase = true;
                        }
                        if (c == parPhrase.charEndInLastPar && i == parPhrase.lastCharParOccurance)
                        {
                            isInEndOfPhrase = true;
                        }

                    }
                    if (isInPhrase == false && isInEndOfPhrase == false)
                    {
                        outStringBuilder.Append(textChars[c]);
                    }
                }
                text.Text = outStringBuilder.ToString();
                i = i + 1;
            }

            return wordprocessingDocument;
        }

        private List<WordMatchedPhrase> FindWordMatchedPhrases(WordprocessingDocument wordprocessingDocument, string replaceWhat)
        {
            char[] replaceWhatChars = replaceWhat.ToCharArray();
            int overlapsRequired = replaceWhatChars.Length;
            int currentChar = 0;
            int firstCharParOccurance = 0;
            int lastCharParOccurance = 0;
            int startChar = 0;
            int endChar = 0;
            List<WordMatchedPhrase> wordMatchedPhrases = new List<WordMatchedPhrase>();
            Document document = wordprocessingDocument.MainDocumentPart.Document;
            int i = 0;
            foreach (Text text in document.Descendants<Text>())
            {
                char[] textChars = text.Text.ToCharArray();
                for (int c = 0; c < textChars.Length; c++)
                {
                    char compareToChar = replaceWhatChars[currentChar];
                    if (textChars[c] == compareToChar)
                    {
                        currentChar = currentChar + 1;
                        if (currentChar == 1)
                        {
                            startChar = c;
                            firstCharParOccurance = i;
                        }
                        if (currentChar == overlapsRequired)
                        {
                            endChar = c;
                            lastCharParOccurance = i;
                            WordMatchedPhrase matchedPhrase = new WordMatchedPhrase
                            {
                                firstCharParOccurance = firstCharParOccurance,
                                lastCharParOccurance = lastCharParOccurance,
                                charEndInLastPar = endChar,
                                charStartInFirstPar = startChar
                            };
                            wordMatchedPhrases.Add(matchedPhrase);
                            currentChar = 0;
                        }
                    }
                    else
                    {
                        currentChar = 0;

                    }
                }
                i = i + 1;
            }

            return wordMatchedPhrases;

        }

        public bool InsertAPicture(WordprocessingDocument wordprocessingDocument, MemoryStream streamImg, string tagText, int widthSize, int heightSize)
        {
            try
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

                imagePart.FeedData(streamImg);
                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), tagText, widthSize, heightSize);
                //wordprocessingDocument.Close();

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        #endregion

        public WordprocessingDocument ReplaceImgInWordDocumennt(WordprocessingDocument wordprocessingDocument, string replaceWhat, Drawing element)
        {
            List<WordMatchedPhrase> matchedPhrases = FindWordMatchedPhrases(wordprocessingDocument, replaceWhat);

            Document document = wordprocessingDocument.MainDocumentPart.Document;
            int i = 0;
            bool isInPhrase = false;
            bool isInEndOfPhrase = false;
            foreach (Text text in document.Descendants<Text>()) // <<< Here
            {
                char[] textChars = text.Text.ToCharArray();
                List<WordMatchedPhrase> curParPhrases = matchedPhrases.FindAll(a => (a.firstCharParOccurance.Equals(i) || a.lastCharParOccurance.Equals(i)));
                StringBuilder outStringBuilder = new StringBuilder();

                for (int c = 0; c < textChars.Length; c++)
                {
                    if (isInEndOfPhrase)
                    {
                        isInPhrase = false;
                        isInEndOfPhrase = false;
                    }

                    foreach (var parPhrase in curParPhrases)
                    {
                        if (c == parPhrase.charStartInFirstPar && i == parPhrase.firstCharParOccurance)
                        {
                            //outStringBuilder.Append(replaceFor);
                            isInPhrase = true;
                            text.Parent.InsertAfter<Drawing>(element, text);
                            text.Remove();
                        }
                        if (c == parPhrase.charEndInLastPar && i == parPhrase.lastCharParOccurance)
                        {
                            isInEndOfPhrase = true;
                        }

                    }
                    if (isInPhrase == false && isInEndOfPhrase == false)
                    {
                        outStringBuilder.Append(textChars[c]);
                    }
                }
                text.Text = outStringBuilder.ToString();
                i = i + 1;
            }

            return wordprocessingDocument;
        }

        public bool InsertAPictureNew(WordprocessingDocument wordprocessingDocument, MemoryStream streamImg, string tagText, int widthSize, int heightSize)
        {
            try
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);

                imagePart.FeedData(streamImg);
                Drawing element = AddImageToBodyNew(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), tagText, widthSize, heightSize);
                //wordprocessingDocument.Close();
                ReplaceImgInWordDocumennt(wordprocessingDocument, tagText, element);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        private Drawing AddImageToBodyNew(WordprocessingDocument wordDoc, string relationshipId, string tagText, int widthSize, int heightSize)
        {
            //Size size = new Size(400, 300);//400,300
            Size size = new Size(widthSize, heightSize);
            Int64Value width = size.Width * 9525;
            Int64Value height = size.Height * 9525;

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = width, Cy = height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = width, Cy = height }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return element;

        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, string tagText, int widthSize, int heightSize)
        {
            //Size size = new Size(400, 300);//400,300
            Size size = new Size(widthSize, heightSize);
            Int64Value width = size.Width * 9525;
            Int64Value height = size.Height * 9525;

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = width, Cy = height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = width, Cy = height }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to body, the element should be in a Run.
            //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
            //Text textPlaceHolder = wordDoc.MainDocumentPart.Document.Body.Descendants<Text>().Where((x) => x.Text == tagText).First();
            // Insert image (the image created with your function) after text place holder.        
            //textPlaceHolder.Parent.InsertAfter<Drawing>(element, textPlaceHolder);
            // Remove text place holder.
            //textPlaceHolder.Remove();
            var document = wordDoc.MainDocumentPart.Document;
            //LambdaLogger.Log("intro addIma/ tag: "+tagText);
            foreach (var text in document.Descendants<Text>()) // <<< Here
            {
                if (text.Text.Contains(tagText))
                {
                    text.Parent.InsertAfter<Drawing>(element, text);
                    text.Remove();
                }
            }
        }

        #region ESTAMPADO DE INFORME DE CONSTRUCCION

        public bool WritingWordReportConstruccionInspection(WordprocessingDocument doc, Inspection inspection)
        {
            try
            {
                //LambdaLogger.Log("entro al WritingWordReportConstruccionInspection");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_DES_RAZONSOCIAL, inspection != null ? inspection.DesRazonsocial : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_DES_DIRECCION, inspection != null ? inspection.DesDireccion : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_FECHA_INSPECCION, inspection != null ? inspection.FecInspeccion : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_NOMBRE_USUARIO_INSPECTOR, inspection != null ? inspection.NombreUsuarioInspector : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_NOMBRE_CONTRATANTE_DES_RAZONSOCIAL, inspection != null ? inspection.DesRazonsocial : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_NUM_LATITUD, inspection != null ? inspection.NumLatitud : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INSPECTION_NUM_LONGITUD, inspection != null ? inspection.NumLongitud : "");
                //_wordDocumentHelper.InsertAPicture(doc, streamImg, ReportConstructionTag.Logo, 400, 300);
                //LambdaLogger.Log("salio del WritingWordReportConstruccionInspection");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public bool WritingWordReportConstruccionInterviewed(WordprocessingDocument doc, Interviewed interviewed)
        {
            try
            {
                //LambdaLogger.Log("entro al WritingWordReportConstruccionInterviewed");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INTERVIEWED_NOM_ENTREVISTADO, interviewed != null ? interviewed.NomEntrevistado : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INTERVIEWED_NOM_CARGO, interviewed != null ? interviewed.DesCargo : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INTERVIEWED_CORREO, interviewed != null ? interviewed.DesCorreo : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_INTERVIEWED_TELEFONO, interviewed != null ? interviewed.DesTelefono : "");
                //_wordDocumentHelper.InsertAPicture(doc, streamImg, ReportConstructionTag.Logo, 400, 300);
                //LambdaLogger.Log("salio del WritingWordReportConstruccionInterviewed");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public bool WritingWordReportConstruccionRisk(WordprocessingDocument doc, Risk risk)
        {
            try
            {
                //LambdaLogger.Log("entro al WritingWordReportConstruccionRisk");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_TERREMOTO, risk != null ? risk.DesTerremoto : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_INUNDACION, risk != null ? risk.DesInundacion : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_LLUVIA, risk != null ? risk.DesLluvia : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_MAREMOTO, risk != null ? risk.DesMaremoto : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_GEOLOGICO, risk != null ? risk.DesGeologico : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_GEOTECNIC, risk != null ? risk.DesGeotecnic : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_EXPOSIC, risk != null ? risk.DesExposic : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_INCENDIO, risk != null ? risk.DesIncendio : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_PERD_BEN, risk != null ? risk.DesPerdBen : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_RISK_DES_RESP_CIVIL, risk != null ? risk.DesRespCivil : "");
                //_wordDocumentHelper.InsertAPicture(doc, streamImg, ReportConstructionTag.Logo, 400, 300);
                //LambdaLogger.Log("salio del WritingWordReportConstruccionRisk");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public bool WritingWordReportConstruccionSinister(WordprocessingDocument doc, Sinister sinister)
        {
            try
            {
                //LambdaLogger.Log("entro al WritingWordReportConstruccionSinister");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_SINISTER_FEC_SINIESTRO, sinister != null ? sinister.FecSiniestro : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_SINISTER_DES_SINIESTRO, sinister != null ? sinister.DesSiniestro : "");
                ReplaceStringInWordDocumennt(doc, ReportConstructionTag.REPORT_CONSTRUCTION_SINISTER_DES_TIE_PARAL, sinister != null ? sinister.DesTieParal : "");
                //_wordDocumentHelper.InsertAPicture(doc, streamImg, ReportConstructionTag.Logo, 400, 300);
                //LambdaLogger.Log("salio del WritingWordReportConstruccionSinister");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        public bool AddImgToWordReportConstruccionArchive(WordprocessingDocument doc, List<DocumentTypeFile> listDocumentTypeFiles)
        {
            try
            {
                if (listDocumentTypeFiles.Count > 0)
                {

                    foreach (var documentTypeFile in listDocumentTypeFiles)
                    {
                        if (documentTypeFile.TypeFile == "01")
                        {
                            InsertAPictureNew(doc, documentTypeFile.File, ReportConstructionTag.REPORT_CONSTRUCTION_ARCHIVE_LOGO, FileSizeDocument.WIDTH_SIZE_LOGO_TYPE_01, FileSizeDocument.HEIGHT_SIZE_LOGO_TYPE_01);
                        }

                        if (documentTypeFile.TypeFile == "02")
                        {
                            InsertAPictureNew(doc, documentTypeFile.File, ReportConstructionTag.REPORT_CONSTRUCTION_ARCHIVE_FIRMA, FileSizeDocument.WIDTH_SIZE_FIRMA_TYPE_02, FileSizeDocument.HEIGHT_SIZE_FIRMA_TYPE_02);
                        }
                    }
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }

        #endregion
    }
}
