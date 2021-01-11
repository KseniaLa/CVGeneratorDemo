using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SelectPdf;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace CVGenerator
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

               try
               {
                    //                    var htmlString = @"<html>
                    // <body>
                    //  <h1>Hello World from selectpdf.com</h1>
                    // </body>
                    //</html>
                    //";

                    //                    HtmlToPdf converter = new HtmlToPdf();

                    //                    // set converter options
                    //                    converter.Options.PdfPageSize = PdfPageSize.A4;
                    //                    converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                    //                    converter.Options.WebPageWidth = 100;
                    //                    converter.Options.WebPageHeight = 100;

                    //                    // create a new pdf document converting an url
                    //                    PdfDocument doc = converter.ConvertHtmlString(htmlString);

                    //                    // save pdf document
                    //                    var bytes = doc.Save();

                    //                    // close pdf document
                    //                    doc.Close();

                    using (MemoryStream mem = new MemoryStream())
                    {
                         using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                         {
                              MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                              // Create the document structure and add some text.
                              mainPart.Document = new Document();
                              Body body = mainPart.Document.AppendChild(new Body());
                              Paragraph para = body.AppendChild(new Paragraph());
                              Run run = para.AppendChild(new Run());
                              run.AppendChild(new Text("Hello world!!!!!!"));
                              mainPart.Document.Save();
                              wordDocument.Close();

                             
                         }
                         mem.Seek(0, SeekOrigin.Begin);
                         return new FileContentResult(mem.ToArray(), "application/octet-stream")
                         {
                              FileDownloadName = "file.docx"
                         };
                    }

                    
               }
               catch(Exception ex)
               {
                    log.LogError(ex.ToString());
                    return new FileContentResult(new byte[] { }, "application/octet-stream")
                    {
                         FileDownloadName = "file.pdf"
                    };
               }
          }
    }
}
