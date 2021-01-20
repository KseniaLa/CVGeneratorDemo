using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.DocumentStylesConfiguration
{
     public static class DocStylesStorage
     {
          public static List<DocStyle> DocStyles = new List<DocStyle>
          {
               new DocStyle
               {
                    Id = 1,
                    Name = "Raw",
                    Style = new StyleConfig
                    {
                         IsStyledHeading = false,
                         Font = "Times New Roman",
                         NameFontSize = 14,
                         HeadingFontSize = 11,
                         TextFontSize = 11
                    }
               },
               new DocStyle
               {
                    Id = 2,
                    Name = "Basic",
                    Style = new StyleConfig
                    {
                         IsStyledHeading = true,
                         Font = "Arial",
                         NameFontSize = 20,
                         HeadingFontSize = 16,
                         TextFontSize = 14
                    }
               },
          };
     }
}
