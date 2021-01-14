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
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

               var styles = new List<DocStyle> 
               {
                    new DocStyle
                    {
                         Id = 1, 
                         Name = "Simple",
                         Style = new StyleConfig
                         {
                              IsStyledHeading = false
                         }
                    },
                    new DocStyle
                    {
                         Id = 2,
                         Name = "Basic",
                         Style = new StyleConfig
                         {
                              IsStyledHeading = true
                         }
                    },
               };


            return new OkObjectResult(styles);
        }
    }
}
