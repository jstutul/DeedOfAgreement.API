using DeedOfAgreement.API.Helpers;
using DeedOfAgreement.API.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;
using HtmlToOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Web.Http;
using System.Web.Http.Cors;
using System.Web.Optimization;
using System.Web.UI.WebControls;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;

namespace DeedOfAgreement.API.Controllers
{
    [EnableCors(origins: "*", headers: "*", methods: "*")]
    [RoutePrefix("api")]
    public class ConverterController : ApiController
    {
        [HttpPost]
        [Route("html-to-word")]
        public HttpResponseMessage ConvertHtmlToWord([FromBody] HtmlRequest db)
        {
            var sqlQuery = $"USP_GET_DEED_OF_AGREEMENT_NEW {db.OrganizationId},{db.ProjectId},'{db.FileNo}'";
            var data = Helper.GetDataFromBackDB(sqlQuery).Tables[0].ToList<HtmlResponse>().FirstOrDefault();

            if (data == null)
                return Request.CreateErrorResponse(HttpStatusCode.NotFound, "No data found");

            var templatePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Templates",
                data.SaleTypeShortName == "I" ? "inst.html" : "ata.html"
            );

            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(templatePath);

            string finalHtml = htmlDoc.DocumentNode.OuterHtml;
            var current = Helper.GetDateInFullWords(DateTime.Now);
            if (data != null)
            {
                finalHtml = finalHtml.Replace("_ProjectName_", data.ProjectName.ToUpper());
                finalHtml = finalHtml.Replace("_Current_", current);
                string phaseSuffix = data.ProjectName == "Purbachal Probashi Palli" ? " phase-1" : "";
                finalHtml = finalHtml.Replace("_ProjectNamePhase_", (data.ProjectName + phaseSuffix).ToUpper());

                finalHtml = finalHtml.Replace("_ClientDetails", data.ClientDetails);
                finalHtml = finalHtml.Replace("_NomineeList_", data.NomineeList);
                finalHtml = finalHtml.Replace("_OptionOne_", data.OptionOne);
                finalHtml = finalHtml.Replace("_OptionTwo_", data.OptionTwo);
                finalHtml = finalHtml.Replace("_OptionA_", data.A);
                finalHtml = finalHtml.Replace("_OptionB_", data.B);
                finalHtml = finalHtml.Replace("_OptionC_", data.C);
                finalHtml = finalHtml.Replace("_Extra_Four_Month_", (data.TotalInstallment + 4).ToString());
                finalHtml = finalHtml.Replace("_OrganizationName_", data.OrganizationName.ToUpper());
                finalHtml = finalHtml.Replace("_BText_", data.Btext);
                finalHtml = finalHtml.Replace("_ScheduleB_", data.ScheduleB);
                finalHtml = finalHtml.Replace(" ,", ",");
                finalHtml = finalHtml.Replace("_forcesmall_", "style='font-size:13.5px;'");
                finalHtml = finalHtml.Replace("LTD.,", "LTD.");
            }
            // 🔁 Replace placeholders

            byte[] fileBytes;

            using (var memStream = new MemoryStream())
            {
                using (var wordDoc = WordprocessingDocument.Create(
                    memStream,
                    WordprocessingDocumentType.Document,
                    true))
                {
                    var mainPart = wordDoc.AddMainDocumentPart();

                    mainPart.Document = new Document(new DocumentFormat.OpenXml.Wordprocessing.Body());

                    AddDefaultStyle(mainPart);

                    var footerPart = AddFooterWithPageNumber(mainPart);
                    string footerPartId = mainPart.GetIdOfPart(footerPart);

                    var sectionProps = new SectionProperties(
                        new FooterReference()
                        {
                            Type = HeaderFooterValues.Default,
                            Id = footerPartId
                        },
                        new PageSize()
                        {
                            Width = 12240U,   // 8.5 inch
                            Height = 20160U   // 14 inch (LEGAL)
                        },
                        new PageMargin
                        {
                            Top = 7920,
                            Bottom = 2016,
                            Left = 1296,
                            Right = 720,
                            Footer = 1152U
                        }
                    );

                    var converter = new HtmlConverter(mainPart);

                    //1st to before last page
                    var splittedText = finalHtml.Split(
                        new string[] { "_page_break_" },
                        StringSplitOptions.None
                    );
                    var firstPageHtml = splittedText[0];
                    var lastPageHtml = splittedText[1];

                    var element_1 = converter.Parse(firstPageHtml);
                    mainPart.Document.Body.Append(element_1);
                    mainPart.Document.Body.Append(
                        new Paragraph(
                            new Run(
                                new Break() { Type = BreakValues.Page }
                            )
                        )
                    );
                    var element_2 = converter.Parse(lastPageHtml);
                    mainPart.Document.Body.Append(element_2);

                    mainPart.Document.Body.Append(sectionProps);
                    mainPart.Document.Save();

                    //var elements = converter.Parse(finalHtml);

                    //mainPart.Document.Body.Append(elements);
                    //mainPart.Document.Body.Append(sectionProps);
                    //mainPart.Document.Save();
                }

                fileBytes = memStream.ToArray();
            }

            // 📦 Return Word file
            var response = Request.CreateResponse(HttpStatusCode.OK);
            response.Content = new ByteArrayContent(fileBytes);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "\"" + db.FileName + "\"" // Added quotes here
            };

            return response;
        }
        string SafeReplace(string html, string placeholder, string value)
        {
            if (string.IsNullOrEmpty(value)) return html.Replace(placeholder, "");
            // This ensures characters like &, <, and > are converted to &amp;, &lt;, etc.
            return html.Replace(placeholder, System.Net.WebUtility.HtmlEncode(value));
        }
        // ================= STYLES =================
        private static FooterPart AddFooterWithPageNumber(MainDocumentPart mainPart)
        {
            FooterPart footerPart = mainPart.AddNewPart<FooterPart>();

            Footer footer = new Footer();

            Paragraph paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Right }
                )
            );

            // Run properties (Pure Black + Normal Size)
            RunProperties runProps = new RunProperties(
                new RunFonts() { Ascii = "Times New Roman", HighAnsi = "Times New Roman" },
                new DocumentFormat.OpenXml.Wordprocessing.FontSize() { Val = "21" }, // 11pt
                new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = "000000" } // PURE BLACK
            );

            // "Page "
            paragraph.Append(new Run(runProps.CloneNode(true), new Text("Page ") { Space = SpaceProcessingModeValues.Preserve }));

            // PAGE field
            paragraph.Append(
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" PAGE ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                new Run(runProps.CloneNode(true), new Text("1")),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
            );

            // " of "
            paragraph.Append(new Run(runProps.CloneNode(true), new Text(" of ") { Space = SpaceProcessingModeValues.Preserve }));

            // NUMPAGES field
            paragraph.Append(
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" NUMPAGES ") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.Separate }),
                new Run(runProps.CloneNode(true), new Text("1")),
                new Run(new FieldChar() { FieldCharType = FieldCharValues.End })
            );

            footer.Append(paragraph);
            footerPart.Footer = footer;
            footerPart.Footer.Save();

            return footerPart;
        }


        private static void AddDefaultStyle(MainDocumentPart mainPart)
        {
            var stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            var styles = new DocumentFormat.OpenXml.Wordprocessing.Styles();

            var docDefaults = new DocDefaults(
                new RunPropertiesDefault(
                    new RunProperties(
                        new RunFonts
                        {
                            Ascii = "Times New Roman",
                            HighAnsi = "Times New Roman",
                            ComplexScript = "SolaimanLipi"
                        },
                        new DocumentFormat.OpenXml.Wordprocessing.FontSize { Val = "21" },
                        new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "000000" }
                    )
                ),
                new ParagraphPropertiesDefault()
            );

            var normalStyle = new DocumentFormat.OpenXml.Wordprocessing.Style
            {
                Type = StyleValues.Paragraph,
                StyleId = "Normal",
                Default = true
            };
            normalStyle.Append(new StyleName { Val = "Normal" });

            styles.Append(docDefaults);
            styles.Append(normalStyle);

            stylePart.Styles = styles;
            stylePart.Styles.Save();
        }
    }

}