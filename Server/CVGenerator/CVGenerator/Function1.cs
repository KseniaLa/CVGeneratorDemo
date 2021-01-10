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
               var htmlString = @"<html>
 <body>
  <h1>Hello World from selectpdf.com</h1>
 </body>
</html>
";

               HtmlToPdf converter = new HtmlToPdf();

               // set converter options
               converter.Options.PdfPageSize = PdfPageSize.A4;
               converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
               converter.Options.WebPageWidth = 100;
               converter.Options.WebPageHeight = 100;

               // create a new pdf document converting an url
               PdfDocument doc = converter.ConvertHtmlString(htmlString);

               // save pdf document
               var bytes = doc.Save();

               // close pdf document
               doc.Close();

               return new FileContentResult(bytes, "application/octet-stream")
               {
                    FileDownloadName = "file.pdf"
               };
          }
    }
}
