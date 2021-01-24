using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Linq;
using CVGenerator.Models;
using CVGenerator.DocumentStylesConfiguration;
using CVGenerator.CvGeneration;

namespace CVGenerator
{
     public static class DocumentGenerationFunction
     {
          private static readonly string _cvFileName = "CV.docx";

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
               try
               {
                    log.LogInformation("Start DocumentGenerationFunction execution");

                    var requestBody = await new StreamReader(req.Body).ReadToEndAsync();

                    var cvRequestModel = JsonConvert.DeserializeObject<CvRequestModel>(requestBody, Settings);

                    var styleId = cvRequestModel.StyleId;
                    var cvInfo = cvRequestModel.CvInfo;

                    var errorItem = ValidateCvRequest(styleId, cvInfo);
                    if (errorItem != null)
                    {
                         log.LogWarning($"Invalid CV request object: {errorItem.ErrorMessage}");
                         return new BadRequestObjectResult(errorItem);
                    }

                    var fileBytes = GenerateCvFileBytes(styleId, cvInfo);

                    return new FileContentResult(fileBytes, "application/octet-stream")
                    {
                         FileDownloadName = _cvFileName
                    };
               }
               catch (Exception ex)
               {
                    log.LogError(ex.ToString());
                    return new BadRequestObjectResult(new ErrorItem
                    {
                         ErrorMessage = $"Fatal error occurred: {ex.Message}"
                    });
               }
               finally
               {
                    log.LogInformation("Finish DocumentGenerationFunction execution");
               }
          }

          private static ErrorItem ValidateCvRequest(int styleId, CvInfoModel cvInfo)
          {
               if (styleId < 0)
               {
                    return new ErrorItem { ErrorMessage = "Invalid CV style Id" };
               }

               if (cvInfo == null)
               {
                    return new ErrorItem { ErrorMessage = "Empty CV info object" };
               }

               if (string.IsNullOrEmpty(cvInfo.FirstName) || string.IsNullOrEmpty(cvInfo.LastName))
               {
                    return new ErrorItem { ErrorMessage = "CV info should have first and last names" };
               }

               return null;
          }

          private static byte[] GenerateCvFileBytes(int styleId, CvInfoModel cvInfo)
          {
               var styles = DocStylesStorage.DocStyles;
               var currStyle = styles.FirstOrDefault(s => s.Id == styleId);

               if (currStyle == null)
               {
                    currStyle = styles.First();
               }

               var cv = CvGenerator.GenerateCvBytes(currStyle.Style, cvInfo);

               return cv;
          }
     }
}
