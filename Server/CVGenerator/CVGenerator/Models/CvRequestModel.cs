using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.Models
{
     public class CvRequestModel
     {
          [JsonProperty("styleId")]
          public int StyleId { get; set; }

          [JsonProperty("cvInfo")]
          public CvInfoModel CvInfo { get; set; }
     }
}
