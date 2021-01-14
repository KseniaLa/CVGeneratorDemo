using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SelectPdf;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Linq;

namespace CVGenerator
{
    public static class DocumentGenerationFunction
    {
        [FunctionName("DocumentGenerationFunction")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

               try
               {
                    using (MemoryStream mem = new MemoryStream())
                    {
                         using (WordprocessingDocument wordDocument =
        WordprocessingDocument.Create(mem, WordprocessingDocumentType.Document, true))
                         {
                              MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                              // Create the document structure and add some text.
                              mainPart.Document = new Document();
                              Body body = mainPart.Document.AppendChild(new Body());

                              Paragraph p = new Paragraph();
                              // Run 1
                              Run r1 = new Run();
                              Text t1 = new Text("Pellentesque ") { Space = SpaceProcessingModeValues.Preserve };
                              // The Space attribute preserve white space before and after your text
                              r1.Append(t1);
                              p.Append(r1);

                              // Run 2 - Bold
                              Run r2 = new Run();
                              RunProperties rp2 = new RunProperties();
                              rp2.Bold = new Bold();
                              // Always add properties first
                              r2.Append(rp2);
                              Text t2 = new Text("commodo ") { Space = SpaceProcessingModeValues.Preserve };
                              r2.Append(t2);
                              p.Append(r2);

                              // Run 3
                              Run r3 = new Run();
                              Text t3 = new Text("rhoncus ") { Space = SpaceProcessingModeValues.Preserve };
                              r3.Append(t3);
                              p.Append(r3);

                              // Run 4 – Italic
                              Run r4 = new Run();
                              RunProperties rp4 = new RunProperties();
                              rp4.Italic = new Italic();
                              // Always add properties first
                              r4.Append(rp4);
                              Text t4 = new Text("mauris") { Space = SpaceProcessingModeValues.Preserve };
                              r4.Append(t4);
                              p.Append(r4);

                              // Run 5
                              Run r5 = new Run();
                              Text t5 = new Text(", sit ") { Space = SpaceProcessingModeValues.Preserve };
                              r5.Append(t5);
                              p.Append(r5);

                              // Run 6 – Italic , bold and underlined
                              Run r6 = new Run();
                              RunProperties rp6 = new RunProperties();
                              rp6.Italic = new Italic();
                              rp6.Bold = new Bold();
                              rp6.Underline = new Underline();
                              // Always add properties first
                              r6.Append(rp6);
                              Text t6 = new Text("amet ") { Space = SpaceProcessingModeValues.Preserve };
                              r6.Append(t6);
                              p.Append(r6);

                              // Run 7
                              Run r7 = new Run();
                              Text t7 = new Text("faucibus arcu ") { Space = SpaceProcessingModeValues.Preserve };
                              r7.Append(t7);
                              p.Append(r7);

                              // Run 8 – Red color
                              Run r8 = new Run();
                              RunProperties rp8 = new RunProperties();
                              rp8.Color = new Color() { Val = "FF0000" };
                              // Always add properties first
                              r8.Append(rp8);
                              Text t8 = new Text("porttitor ") { Space = SpaceProcessingModeValues.Preserve };
                              r8.Append(t8);
                              p.Append(r8);

                              // Run 9
                              Run r9 = new Run();
                              Text t9 = new Text("pharetra. Maecenas quis erat quis eros iaculis placerat ut at mauris.") { Space = SpaceProcessingModeValues.Preserve };
                              r9.Append(t9);
                              p.Append(r9);
                              // Add your paragraph to docx body
                              body.Append(p);

                              //

                              // Heading 1
                              StyleRunProperties styleRunPropertiesH1 = new StyleRunProperties();
                              Color color1 = new Color() { Val = "2F5496" };
                              // Specify a 16 point size. 16x2 because it’s half-point size
                              DocumentFormat.OpenXml.Wordprocessing.FontSize fontSize1 = new DocumentFormat.OpenXml.Wordprocessing.FontSize();
                              fontSize1.Val = new StringValue("32");

                              styleRunPropertiesH1.Append(color1);
                              styleRunPropertiesH1.Append(fontSize1);
                              // Check above at the begining of the word creation to check where mainPart come from
                              AddStyleToDoc(mainPart, "heading1", "heading 1", styleRunPropertiesH1);

                              //// Heading 2
                              //StyleRunProperties styleRunPropertiesH2 = new StyleRunProperties();
                              //Color color2 = new Color() { Val = "2F5496" };
                              //// Specify a 13 point size. 16x2 because it’s half-point size
                              //DocumentFormat.OpenXml.Wordprocessing.FontSize fontSize1 = new DocumentFormat.OpenXml.Wordprocessing.FontSize();
                              //fontSize1.Val = new StringValue("26");

                              //styleRunPropertiesH1.Append(color1);
                              //styleRunPropertiesH1.Append(fontSize1);
                              //AddStyleToDoc(mainPart, "heading2", "heading 2", styleRunPropertiesH1);

                              // Create your heading1 into docx
                              Paragraph pH1 = new Paragraph();
                              ParagraphProperties ppH1 = new ParagraphProperties();
                              ppH1.ParagraphStyleId = new ParagraphStyleId() { Val = "heading1" };
                              ppH1.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" };
                              pH1.Append(ppH1);
                              // Run Heading1
                              Run rH1 = new Run();
                              Text tH1 = new Text("First Heading") { Space = SpaceProcessingModeValues.Preserve };
                              rH1.Append(tH1);
                              pH1.Append(rH1);
                              // Add your heading to docx body
                              body.Append(pH1);

                              // Create your heading2 into docx
                              Paragraph pH2 = new Paragraph();
                              ParagraphProperties ppH2 = new ParagraphProperties();
                              ppH2.ParagraphStyleId = new ParagraphStyleId() { Val = "heading2" };
                              ppH2.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" };
                              pH2.Append(ppH2);
                              // Run Heading2
                              Run rH2 = new Run();
                              Text tH2 = new Text("Second Heading") { Space = SpaceProcessingModeValues.Preserve };
                              rH2.Append(tH2);
                              pH2.Append(rH2);
                              // Add your heading2 to docx body
                              body.Append(pH2);

                              //body.Append(new Paragraph(new Run(new Text("Lorem ipsum dolor sit amet, consectetur adipiscing elit. Praesent quam augue, tempus id metus in, laoreet viverra quam. Sed vulputate risus lacus, et dapibus orci porttitor non."))));
                              mainPart.Document.Save();
                              wordDocument.Close();

                             
                         }
                         mem.Seek(0, SeekOrigin.Begin);
                         return new FileContentResult(mem.ToArray(), "application/octet-stream")
                         {
                              FileDownloadName = "file.docx"
                         };
                    }

                    
               }
               catch(Exception ex)
               {
                    log.LogError(ex.ToString());
                    return new FileContentResult(new byte[] { }, "application/octet-stream")
                    {
                         FileDownloadName = "file.pdf"
                    };
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
