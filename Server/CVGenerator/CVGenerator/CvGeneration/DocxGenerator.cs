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
          public static byte[] GenerateFileBytes(List<TextItem> textItems, string font, bool hasStyledHeadings, int headingFontSize)
          {
               using (var mem = new MemoryStream())
               {
                    using (var wordDocument = WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                    {
                         var mainPart = wordDocument.AddMainDocumentPart();

                         mainPart.Document = new Document();
                         var body = mainPart.Document.AppendChild(new Body());

                         if (hasStyledHeadings)
                         {
                              var styleRunPropertiesH1 = new StyleRunProperties();
                              var color1 = new Color() { Val = "2F5496" };

                              var fontSize1 = new FontSize
                              {
                                   Val = new StringValue((headingFontSize * 2).ToString())
                              };

                              styleRunPropertiesH1.Append(color1);
                              styleRunPropertiesH1.Append(fontSize1);
                              // Check above at the begining of the word creation to check where mainPart come from
                              AddStyleToDoc(mainPart, "heading1", "heading 1", styleRunPropertiesH1);
                         }

                         foreach (var textItem in textItems)
                         {
                              var p = new Paragraph();

                              var pProps = new ParagraphProperties
                              {
                                   Justification = new Justification() { Val = JustificationValues.Start },
                                   ParagraphStyleId = textItem.IsStyledHeading ? new ParagraphStyleId() { Val = "heading1" } : null,
                                   SpacingBetweenLines = new SpacingBetweenLines() { After = (textItem.IsHeading ? "0" : "200") },
                              };
                              p.Append(pProps);

                              var r1 = new Run();
                              RunProperties rPr = new RunProperties(
                               new RunFonts()
                               {
                                    Ascii = font
                               });

                              r1.Append(rPr);

                              if (textItem.IsStyledHeading)
                              {
                                   //ParagraphProperties ppH1 = new ParagraphProperties
                                   //{
                                   //     ParagraphStyleId = new ParagraphStyleId() { Val = "heading1" },
                                   //     SpacingBetweenLines = new SpacingBetweenLines() { After = "0" },
                                   //};
                                   //p.Append(ppH1);
                                   Text tH1 = new Text(textItem.Text) { Space = SpaceProcessingModeValues.Preserve };
                                   r1.Append(tH1);
                              }
                              else
                              {
                                   var rp1 = new RunProperties
                                   {
                                        FontSize = new FontSize { Val = new StringValue((textItem.FontSize * 2).ToString()) },
                                        Bold = textItem.IsBold ? new Bold() : null
                                   };

                                   r1.Append(rp1);

                                   Text t1 = new Text(textItem.Text) { Space = SpaceProcessingModeValues.Preserve };
                                   r1.Append(t1);
                              }

                              p.Append(r1);
                              body.Append(p);
                         }

                         mainPart.Document.Save();
                         wordDocument.Close();
                    }

                    mem.Seek(0, SeekOrigin.Begin);
                    return mem.ToArray();
               }
          }

          // Apply a style to a paragraph.
          public static void AddStyleToDoc(MainDocumentPart mainPart, string styleid, string stylename, StyleRunProperties styleRunProperties)
          {

               // Get the Styles part for this document.
               StyleDefinitionsPart part =
                   mainPart.StyleDefinitionsPart;

               // If the Styles part does not exist, add it and then add the style.
               if (part == null)
               {
                    part = AddStylesPartToPackage(mainPart);
                    AddNewStyle(part, styleid, stylename, styleRunProperties);
               }
               else
               {
                    // If the style is not in the document, add it.
                    if (IsStyleIdInDocument(mainPart, styleid) != true)
                    {
                         // No match on styleid, so let's try style name.
                         string styleidFromName = GetStyleIdFromStyleName(mainPart, stylename);
                         if (styleidFromName == null)
                         {
                              AddNewStyle(part, styleid, stylename, styleRunProperties);
                         }
                         else
                              styleid = styleidFromName;
                    }
               }

          }

          // Add a StylesDefinitionsPart to the document.  Returns a reference to it.
          public static StyleDefinitionsPart AddStylesPartToPackage(MainDocumentPart mainPart)
          {
               StyleDefinitionsPart part;
               part = mainPart.AddNewPart<StyleDefinitionsPart>();
               DocumentFormat.OpenXml.Wordprocessing.Styles root = new DocumentFormat.OpenXml.Wordprocessing.Styles();
               root.Save(part);
               return part;
          }

          public static bool IsStyleIdInDocument(MainDocumentPart mainPart, string styleid)
          {
               // Get access to the Styles element for this document.
               DocumentFormat.OpenXml.Wordprocessing.Styles s = mainPart.StyleDefinitionsPart.Styles;

               // Check that there are styles and how many.
               int n = s.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>().Count();
               if (n == 0)
                    return false;

               // Look for a match on styleid.
               DocumentFormat.OpenXml.Wordprocessing.Style style = s.Elements<DocumentFormat.OpenXml.Wordprocessing.Style>()
                   .Where(st => (st.StyleId == styleid) && (st.Type == StyleValues.Paragraph))
                   .FirstOrDefault();
               if (style == null)
                    return false;

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

          // Create a new style with the specified styleid and stylename and add it to the specified style definitions part.
          private static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, string styleid, string stylename, StyleRunProperties styleRunProperties)
          {
               // Get access to the root element of the styles part.
               DocumentFormat.OpenXml.Wordprocessing.Styles styles = styleDefinitionsPart.Styles;

               // Create a new paragraph style and specify some of the properties.
               DocumentFormat.OpenXml.Wordprocessing.Style style = new DocumentFormat.OpenXml.Wordprocessing.Style()
               {
                    Type = StyleValues.Paragraph,
                    StyleId = styleid,
                    CustomStyle = true
               };
               style.Append(new StyleName() { Val = stylename });
               style.Append(new BasedOn() { Val = "Normal" });
               style.Append(new NextParagraphStyle() { Val = "Normal" });
               style.Append(new UIPriority() { Val = 900 });

               // Create the StyleRunProperties object and specify some of the run properties.


               // Add the run properties to the style.
               // --- Here we use the OuterXml. If you are using the same var twice, you will get an error. So to be sure just insert the xml and you will get through the error.
               style.Append(styleRunProperties);

               // Add the style to the styles part.
               styles.Append(style);
          }
     }
}
