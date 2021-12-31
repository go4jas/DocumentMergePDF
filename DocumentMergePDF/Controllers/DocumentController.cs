using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MailMerge;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Serilog;
using Serilog.Extensions.Logging;
using Spire.Doc;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;

namespace DocumentMergePDF.Controllers
{
    public class DocumentController : Controller
    {
        
        [HttpPost("renderdocument")]
        public async Task<IActionResult> Render(List<string> filedNames, List<string> fieldValues, IFormFile template)
        {

            var log = new LoggerConfiguration()
                        .WriteTo.File("logs.txt")
                        .CreateLogger();


            var microsoftLogger = new SerilogLoggerFactory(log)
                .CreateLogger("logger");

            log.Information("render process started");

            
            var fields = new Dictionary<string, string>();
            var index = 0;
            foreach(var item in filedNames)
            {
                fields.Add(item, fieldValues[index]);
                index++;
            }
            var (outputStream, errors) = new MailMerger().Merge(template.OpenReadStream(), fields);

            var obj = new MailMerger(microsoftLogger);
            
            var (results, err) = obj.Merge(template.OpenReadStream(), fields, "out3.docx");

           
            log.Information("render process ended");

            outputStream.Seek(0, SeekOrigin.Begin);

            return File(outputStream, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "outputfile.docx");

        }

        [HttpPost("renderdocumentv2")]
        public async Task<IActionResult> Renderv2(List<string> filedNames, List<string> fieldValues, IFormFile template)
        {

            var log = new LoggerConfiguration()
                        .WriteTo.Console()
                        .CreateLogger();


            var microsoftLogger = new SerilogLoggerFactory(log)
                .CreateLogger("logger");

            log.Information("render process started");

            var fields = new Dictionary<string, string>();
            var index = 0;
            foreach (var item in filedNames)
            {
                fields.Add(item, fieldValues[index]);
                index++;
            }

            var outputStream = new MemoryStream();
            template.OpenReadStream().CopyTo(outputStream);

            ApplyAllKnownMergeTransformationsToMainDocumentPart(fields, outputStream, microsoftLogger);

            /* using (var fileStream = System.IO.File.Create("out.docx"))
             {
                 outputStream.Seek(0, SeekOrigin.Begin);
                 outputStream.CopyTo(fileStream);
             }*/

            Spire.Doc.Document document = new Spire.Doc.Document(outputStream, FileFormat.Docx);
            var pdfoutputStream = new MemoryStream();


            var fontFile = "/opt/app-root/src/DocumentMergePDF/fonts/arial.ttf";
            List<PrivateFontPath> fonts = new List<PrivateFontPath>();
            fonts.Add(new PrivateFontPath("Arial", fontFile));
            fonts.Add(new PrivateFontPath("arial", fontFile));

            ToPdfParameterList ps = new ToPdfParameterList
            {
                UsePSCoversion = true,
                PdfConformanceLevel = Spire.Pdf.PdfConformanceLevel.Pdf_X1A2001,
                PrivateFontPaths = fonts
            };

            document.SaveToStream(pdfoutputStream, ps);

            pdfoutputStream.Seek(0, SeekOrigin.Begin);

            log.Information("render process ended");

            return File(pdfoutputStream, "application/pdf", "outputfile.pdf");

        }


        [HttpPost("renderdocumentv3")]
        public async Task<IActionResult> Renderv3(List<string> filedNames, List<string> fieldValues, IFormFile template)
        {

            var log = new LoggerConfiguration()
                        .WriteTo.Console()
                        .CreateLogger();


            var microsoftLogger = new SerilogLoggerFactory(log)
                .CreateLogger("logger");

            log.Information("render process started");

            var fields = new Dictionary<string, string>();
            var index = 0;
            foreach (var item in filedNames)
            {
                fields.Add(item, fieldValues[index]);
                index++;
            }

            var outputStream = new MemoryStream();
            template.OpenReadStream().CopyTo(outputStream);

            ApplyAllKnownMergeTransformationsToMainDocumentPart(fields, outputStream, microsoftLogger);

            /* using (var fileStream = System.IO.File.Create("out.docx"))
             {
                 outputStream.Seek(0, SeekOrigin.Begin);
                 outputStream.CopyTo(fileStream);
             }*/

            Spire.Doc.Document document = new Spire.Doc.Document(outputStream, FileFormat.Docx);
            var pdfoutputStream = new MemoryStream();

            var images = document.SaveToImages(Spire.Doc.Documents.ImageType.Bitmap);


            ToPdfParameterList ps = new ToPdfParameterList
            {
                UsePSCoversion = true
                //PdfConformanceLevel = Spire.Pdf.PdfConformanceLevel.Pdf_X1A2001,


            };

            document.SaveToStream(pdfoutputStream, ps);

            pdfoutputStream.Seek(0, SeekOrigin.Begin);

            log.Information("render process ended");

            return File(pdfoutputStream, "application/pdf", "outputfile.pdf");

        }


        internal void ApplyAllKnownMergeTransformationsToMainDocumentPart(Dictionary<string, string> fieldValues, Stream workingStream, Microsoft.Extensions.Logging.ILogger microsoftLogger = null)
        {
            var xdoc = GetMainDocumentPartXml(workingStream);

            xdoc.SimpleMergeFields(fieldValues, microsoftLogger);
            xdoc.ComplexMergeFields(fieldValues, microsoftLogger);
            //xdoc.MergeDate(microsoftLogger, DateTime, fieldValues.ContainsKey(DATEKey) ? fieldValues[DATEKey] : DateTime?.ToLongDateString());

            using (var wpDocx = WordprocessingDocument.Open(workingStream, true))
            {
                var bodyNode = xdoc.SelectSingleNode("/w:document/w:body", OoXmlNamespace.Manager);
                var documentBody = new DocumentFormat.OpenXml.Wordprocessing.Body(bodyNode.OuterXml);
                wpDocx.MainDocumentPart.Document.Body = documentBody;
            }
        }

        public XmlDocument GetMainDocumentPartXml(Stream docxStream)
        {
            var xdoc = new XmlDocument(OoXmlNamespace.Manager.NameTable);
            using (var wpDocx = WordprocessingDocument.Open(docxStream, false))
            using (var docOutStream = wpDocx.MainDocumentPart.GetStream(FileMode.Open, FileAccess.Read))
            {
                xdoc.Load(docOutStream);
            }
            return xdoc;
        }

        [HttpGet("document")]
        public async Task<IActionResult> Get(string contraventionNumber, string documentId, string documentName)
        {
            /* var fileContent = System.IO.File.ReadAllText(SRC);
             fileContent = fileContent.Replace("�ContraventionNumber�", "A123456789B");
             fileContent = fileContent.Replace("�TrafficSafetyActSection�", "TSA 60(a)");
             fileContent = fileContent.Replace("�TrafficSafetyActDescription�", "Driving without proper insurance");
             fileContent = fileContent.Replace("�OccurrenceTime�", "11:10");
             fileContent = fileContent.Replace("�OccurrenceDate�", "2022-02-12");
             fileContent = fileContent.Replace("�FineAmount�", "$200");
             fileContent = fileContent.Replace("�Demerits�", "3");
             fileContent = fileContent.Replace("�RecipientName�", "John Doh");
             fileContent = fileContent.Replace("�LicencePlate�", "AB123O");
             fileContent = fileContent.Replace("�IssueDate�", "2022-02-12");

            //select html to pdf
             HtmlToPdf converter = new HtmlToPdf();
             PdfDocument doc = converter.ConvertHtmlString(fileContent);//converter.ConvertUrl(url);
             doc.Save(DEST);
             doc.Close();

             FileStream fileStream = new FileStream(DEST, FileMode.Open);
             return new FileStreamResult(fileStream, "application/pdf");

             //return this.Ok(documentName);*/
            throw new NotImplementedException();
        }
    }
}
