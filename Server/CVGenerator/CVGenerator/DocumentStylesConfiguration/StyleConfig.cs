using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.DocumentStylesConfiguration
{
     public class StyleConfig
     {
          public bool IsStyledHeading {get; set;}
          public string Font { get; set; }
          public int NameFontSize { get; set; }
          public int HeadingFontSize { get; set; }
          public int TextFontSize { get; set; }
     }
}
