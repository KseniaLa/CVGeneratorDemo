using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CVGenerator.CvGeneration
{
     public static class DocxGenerator
     {
          private static readonly string _headingFontColor = "2F5496";

          public static byte[] GenerateFileBytes(List<TextItem> textItems, string font, bool hasStyledHeadings, int headingFontSize)
          {
               using var memoryStream = new MemoryStream();

               using (var wordDocument = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
               {
                    var mainPart = wordDocument.AddMainDocumentPart();

                    mainPart.Document = new Document();
                    var body = mainPart.Document.AppendChild(new Body());

                    if (hasStyledHeadings)
                    {
                         var styleRunProperties = new StyleRunProperties();
                         var fontColor = new Color() { Val = _headingFontColor };

                         var fontSize = new FontSize
                         {
                              Val = new StringValue((headingFontSize * 2).ToString())
                         };

                         styleRunProperties.Append(fontColor);
                         styleRunProperties.Append(fontSize);
                         AddStyleToDoc(mainPart, "heading1", "heading 1", styleRunProperties);
                    }

                    foreach (var textItem in textItems)
                    {
                         var paragraph = new Paragraph();

                         var pProps = new ParagraphProperties
                         {
                              Justification = new Justification() { Val = JustificationValues.Start },
                              ParagraphStyleId = textItem.IsStyledHeading ? new ParagraphStyleId() { Val = "heading1" } : null,
                              SpacingBetweenLines = new SpacingBetweenLines() { After = (textItem.IsHeading ? "0" : "200") },
                         };
                         paragraph.Append(pProps);

                         var run = new Run();
                         var runPropsBase = new RunProperties(new RunFonts() { Ascii = font });

                         run.Append(runPropsBase);

                         if (textItem.IsStyledHeading)
                         {
                              var headingText = new Text(textItem.Text) { Space = SpaceProcessingModeValues.Preserve };
                              run.Append(headingText);
                         }
                         else
                         {
                              var runProps = new RunProperties
                              {
                                   FontSize = new FontSize { Val = new StringValue((textItem.FontSize * 2).ToString()) },
                                   Bold = textItem.IsBold ? new Bold() : null
                              };

                              run.Append(runProps);

                              var text = new Text(textItem.Text) { Space = SpaceProcessingModeValues.Preserve };
                              run.Append(text);
                         }

                         paragraph.Append(run);
                         body.Append(paragraph);
                    }

                    mainPart.Document.Save();
                    wordDocument.Close();
               }

               memoryStream.Seek(0, SeekOrigin.Begin);
               return memoryStream.ToArray();
          }

          public static void AddStyleToDoc(MainDocumentPart mainPart, string styleid, string stylename, StyleRunProperties styleRunProperties)
          {
               var stylePart = mainPart.StyleDefinitionsPart;

               if (stylePart == null)
               {
                    stylePart = AddStylesPartToPackage(mainPart);
                    AddNewStyle(stylePart, styleid, stylename, styleRunProperties);
               }
               else
               {
                    if (IsStyleIdInDocument(mainPart, styleid) != true)
                    {
                         string styleidFromName = GetStyleIdFromStyleName(mainPart, stylename);
                         if (styleidFromName == null)
                         {
                              AddNewStyle(stylePart, styleid, stylename, styleRunProperties);
                         }
                    }
               }
          }

          public static StyleDefinitionsPart AddStylesPartToPackage(MainDocumentPart mainPart)
          {
               StyleDefinitionsPart part;
               part = mainPart.AddNewPart<StyleDefinitionsPart>();
               var root = new Styles();
               root.Save(part);
               return part;
          }

          public static bool IsStyleIdInDocument(MainDocumentPart mainPart, string styleid)
          {
               var stypesPart = mainPart.StyleDefinitionsPart.Styles;

               var elemCount = stypesPart.Elements<Style>().Count();
               if (elemCount == 0) return false;

               var style = stypesPart.Elements<Style>()
                   .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                   .FirstOrDefault();
               
               if (style == null) return false;

               return true;
          }

          public static string GetStyleIdFromStyleName(MainDocumentPart doc, string styleName)
          {
               StyleDefinitionsPart stylePart = doc.StyleDefinitionsPart;
               string styleId = stylePart.Styles.Descendants<StyleName>()
                   .Where(s => s.Val.Value.Equals(styleName) &&
                       (((Style)s.Parent).Type == StyleValues.Paragraph))
                   .Select(n => ((Style)n.Parent).StyleId).FirstOrDefault();
               return styleId;
          }

          private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename, StyleRunProperties styleRunProperties)
          {
               var styles = styleDefinitionsPart.Styles;

               var style = new Style()
               {
                    Type = StyleValues.Paragraph,
                    StyleId = styleid,
                    CustomStyle = true
               };
               style.Append(new StyleName() { Val = stylename });
               style.Append(new BasedOn() { Val = "Normal" });
               style.Append(new NextParagraphStyle() { Val = "Normal" });
               style.Append(new UIPriority() { Val = 900 });

               
               style.Append(styleRunProperties);

               styles.Append(style);
          }
     }
}
