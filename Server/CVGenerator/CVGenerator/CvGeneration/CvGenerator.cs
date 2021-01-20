using CVGenerator.DocumentStylesConfiguration;
using CVGenerator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace CVGenerator.CvGeneration
{
     public static class CvGenerator
     {
          public static byte[] GenerateCvBytes(StyleConfig docStyle, CvModel cvModel)
          {
               var textItems = new List<TextItem>();

               if (cvModel.FirstName != null || cvModel.LastName != null)
               {
                    textItems.Add(new TextItem
                    {
                         Text = $"{cvModel.FirstName} {cvModel.LastName}",
                         FontSize = docStyle.NameFontSize,
                         IsHeading = false,
                         IsBold = true
                    });
               }

               if (cvModel.Education != null)
               {
                    textItems.Add(new TextItem
                    {
                         Text = "Education",
                         FontSize = docStyle.HeadingFontSize,
                         IsBold = true,
                         IsHeading = true,
                         IsStyledHeading = docStyle.IsStyledHeading
                    });

                    textItems.Add(new TextItem
                    {
                         Text = cvModel.Education,
                         FontSize = docStyle.TextFontSize,
                         IsBold = false,
                         IsHeading = false,
                         IsStyledHeading = false
                    });
               }

               if (cvModel.WorkingExperience != null)
               {
                    textItems.Add(new TextItem
                    {
                         Text = "Working Experience",
                         FontSize = docStyle.HeadingFontSize,
                         IsBold = true,
                         IsHeading = true,
                         IsStyledHeading = docStyle.IsStyledHeading
                    });

                    textItems.Add(new TextItem
                    {
                         Text = cvModel.WorkingExperience,
                         FontSize = docStyle.TextFontSize,
                         IsBold = false,
                         IsHeading = false,
                         IsStyledHeading = false
                    });
               }

               var bytes = DocxGenerator.GenerateFileBytes(textItems, docStyle.Font, docStyle.IsStyledHeading, docStyle.HeadingFontSize);

               return bytes;
          }
     }
}
