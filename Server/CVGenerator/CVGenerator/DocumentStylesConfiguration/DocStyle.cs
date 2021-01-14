using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.DocumentStylesConfiguration
{
     public class DocStyle
     {
          public int Id { get; set; }
          public string Name { get; set; }

          public StyleConfig Style { get; set; }
     }
}
