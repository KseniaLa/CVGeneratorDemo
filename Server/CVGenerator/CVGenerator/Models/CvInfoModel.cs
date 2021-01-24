using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.Models
{
     public class CvInfoModel
     {
          [JsonProperty("firstName")]
          public string FirstName { get; set; }

          [JsonProperty("lastName")]
          public string LastName { get; set; }

          [JsonProperty("email")]
          public string Email { get; set; }

          [JsonProperty("dateOfBirth")]
          public DateTime DateOfBirth { get; set; }

          [JsonProperty("education")]
          public string Education { get; set; }

          [JsonProperty("workingExperience")]
          public string WorkingExperience { get; set; }
     }
}
