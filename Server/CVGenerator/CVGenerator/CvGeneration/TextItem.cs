using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.CvGeneration
{
     public class TextItem
     {
          public string Text { get; set; }
          public int FontSize { get; set; }
          public bool IsHeading { get; set; }
          public bool IsStyledHeading { get; set; }
          public string CustomStyleName { get; set; }
          public bool IsBold { get; set; }
     }
}
