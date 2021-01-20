using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.Models
{
     public class CvModel
     {
          [JsonProperty("firstName")]
          public string FirstName { get; set; }
          [JsonProperty("lastName")]
          public string LastName { get; set; }
          public string DateOfBirth { get; set; }
          [JsonProperty("education")]
          public string Education { get; set; }
          [JsonProperty("workingExperience")]
          public string WorkingExperience { get; set; }
     }
}
