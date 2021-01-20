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
using System.Linq;
using CVGenerator.Models;
using CVGenerator.DocumentStylesConfiguration;
using CVGenerator.CvGeneration;

namespace CVGenerator
{
     public static class DocumentGenerationFunction
     {
          private static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
          {
               NullValueHandling = NullValueHandling.Ignore,
               MissingMemberHandling = MissingMemberHandling.Ignore,
               DateTimeZoneHandling = DateTimeZoneHandling.Utc
          };

          [FunctionName("DocumentGenerationFunction")]
          public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
          {
               log.LogInformation("C# HTTP trigger function processed a request.");

               try
               {
                    var requestBody = await new StreamReader(req.Body).ReadToEndAsync();

                    var cvModel = JsonConvert.DeserializeObject<CvModel>(requestBody, Settings);

                    var styles = DocStylesStorage.DocStyles;

                    var currStyle = styles.FirstOrDefault(s => s.Id == 1);

                    var cv = CvGenerator.GenerateCvBytes(currStyle.Style, cvModel);

                    return new FileContentResult(cv, "application/octet-stream")
                    {
                         FileDownloadName = "file.docx"
                    };
               }
               catch (Exception ex)
               {
                    log.LogError(ex.ToString());
                    return new FileContentResult(new byte[] { }, "application/octet-stream")
                    {
                         FileDownloadName = "file.docx"
                    };
               }
          }
     }
}
