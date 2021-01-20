using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using CVGenerator.DocumentStylesConfiguration;
using System.Collections.Generic;

namespace CVGenerator
{
     public static class DocumentStylesFunction
     {
          [FunctionName("DocumentStylesFunction")]
          public static IActionResult Run(
              [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
              ILogger log)
          {
               try
               {
                    log.LogInformation("Start DocumentStylesFunction execution");

                    var styles = DocStylesStorage.DocStyles;

                    return new OkObjectResult(styles);
               }
               catch (Exception ex)
               {
                    log.LogError($"Fatal error occurred: {ex}");

                    return new BadRequestObjectResult(ex.Message);
               }
               finally
               {
                    log.LogInformation("Finish DocumentStylesFunction execution");
               }
          }
     }
}
