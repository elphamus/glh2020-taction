using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using taction.DTO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Hosting;

namespace taction.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class TenantController : ControllerBase
    {
        private readonly ILogger<TenantController> _logger;
        private IWebHostEnvironment _hostingEnvironment;

        public TenantController(ILogger<TenantController> logger, IWebHostEnvironment environment)
        {
            _logger = logger;
            _hostingEnvironment = environment;
        }

        [HttpGet]
        public IEnumerable<Tenant> Get()
        {

            return Enumerable.Range(1, 1).Select(i => new Tenant
            {
                name = "Test Name 1",
            }).ToArray();
        }

        [HttpPost]
        public IActionResult Post([FromBody] DataDTO data)
        {
            var fullFilePath = Path.Combine(_hostingEnvironment.ContentRootPath, "Letter.docx");

            // Copy file content to MemeoryStream via byte array
            MemoryStream stream = new MemoryStream();
            byte[] fileBytesArray = System.IO.File.ReadAllBytes(fullFilePath);
            stream.Write(fileBytesArray, 0, fileBytesArray.Length);             // copy file content to MemoryStream
            stream.Position = 0;

            // Edit word document content
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(stream, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;
                var paras = body.Elements<Paragraph>();
                var newParas = new List<Paragraph>();
                // We're loosing the address so we should put it back in.

                newParas.Add(new Paragraph(new ParagraphProperties(new Bold()), new Run(new RunProperties(new Bold()), new Text("Private and Confidential"))));
                newParas.Add(new Paragraph(new Run(new Text(" "))));
                newParas.Add(new Paragraph(new Run(new Text(data.landlord.name))));
                newParas.Add(new Paragraph(new Run(new Text(data.landlord.address1))));
                newParas.Add(new Paragraph(new Run(new Text(data.landlord.address2))));
                if (!String.IsNullOrEmpty(data.landlord.address3))
                {
                    newParas.Add(new Paragraph(new Run(new Text(data.landlord.address3))));
                }
                newParas.Add(new Paragraph(new Run(new Text(data.landlord.city))));
                newParas.Add(new Paragraph(new Run(new Text(data.landlord.postcode))));
                newParas.Add(new Paragraph(new Run(new Text(" "))));

                foreach (var para in paras)
                {
                    if (para.InnerText.Contains(@"[INSERTDATE]"))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[INSERTDATE]", DateTime.UtcNow.Date.ToString("dd/MM/yyyy")));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"[TENANTNAME]"))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[TENANTNAME]", data.tenant.name));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"[TENANTNAMEADDRESSOFPROPERTY]"))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[TENANTNAMEADDRESSOFPROPERTY]", data.tenant.name + " " + data.tenant.address1));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"[Insert description provided by tenant when they summarise the damage.]"))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[Insert description provided by tenant when they summarise the damage.]", data.issue.summary));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"[Insert description provided by tenant when they summarise the effect]."))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[Insert description provided by tenant when they summarise the effect].", data.issue.effects));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"[Name of Landlord]"))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[Name of Landlord]", data.landlord.name));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"[TENANTNAME]"))
                    {
                        Text t = new Text(para.InnerText.Replace(@"[TENANTNAME]", data.tenant.name));
                        para.RemoveAllChildren<Run>();
                        para.AppendChild<Run>(new Run(t));
                    }
                    if (para.InnerText.Contains(@"DD/MM/YYYY – I contacted my landlord via email.]"))
                    {
                        para.RemoveAllChildren<Run>();
                        foreach (var h in data.history)
                        {
                            var element =
                            new Paragraph(
                                new ParagraphProperties(
                                    new ParagraphStyleId() { Val = "ListParagraph" },
                                    new NumberingProperties(
                                    new NumberingLevelReference() { Val = 0 },
                                    new NumberingId() { Val = 1 })),
                                    new Run(new Text(h.date.ToString("dd/MM/yyyy") + " - " + h.description + "\n")))
                            { RsidParagraphAddition = "00031711", RsidParagraphProperties = "00031711", RsidRunAdditionDefault = "00031711" };
                            newParas.Add(element);
                        }
                    }
                    if (para.InnerText.Contains(@"[DD/MM/YYY between 00:00 and 00:00 (ie 24 hour clock)]"))
                    {
                        para.RemoveAllChildren<Run>();
                        foreach (var h in data.availability)
                        {
                            var element =
                            new Paragraph(
                                new ParagraphProperties(
                                    new ParagraphStyleId() { Val = "ListParagraph" },
                                    new NumberingProperties(
                                    new NumberingLevelReference() { Val = 2 },
                                    new NumberingId() { Val = 1 })),
                                    new Run(new Text(h.startAt.ToString("dd/MM/yyyy h:mm tt") + " - " + h.endAt.ToString("h:mm tt") + "\n")))
                            { RsidParagraphAddition = "00031711", RsidParagraphProperties = "00031711", RsidRunAdditionDefault = "00031711" };
                            newParas.Add(element);
                        }
                    }
                    newParas.Add(para);
                }
                body.RemoveAllChildren();
                foreach (var p in newParas)
                {
                    body.AppendChild<Paragraph>(p);
                }

                body.AppendChild<Paragraph>(AddPageBreak());
                //body.AppendChild<Paragraph>(AddLandscapeParagraph());
                body.AppendChild<Table>(CreateScheduleOfConditions(data));

                wordDoc.MainDocumentPart.Document.Save();

            }
            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "servedFilename.docx");

        }

        public Table CreateScheduleOfConditions(DataDTO data)
        {
            /*
             * Border size of 1pt is Size 8, 3pt is Size 24
             */

            // Create an empty table.
            Table table = new Table();

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 8
                    },
                    new BottomBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 8
                    },
                    new LeftBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 8
                    },
                    new RightBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 8
                    },
                    new InsideHorizontalBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 8
                    },
                    new InsideVerticalBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 8
                    }
                )
            );

            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tblProp);

            table.Append(CreateTableHeaders());

            foreach (var d in data.defects)
            {
                TableRow tr = new TableRow();

                TableCell tc0 = new TableCell();
                tc0.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
                tc0.Append(new Paragraph(new Run(new Text(d.item))));
                tr.Append(tc0);

                TableCell tc1 = new TableCell();
                tc1.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
                tc1.Append(new Paragraph(new Run(new Text(d.room))));
                tr.Append(tc1);


                TableCell tc2 = new TableCell();
                tc2.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
                tc2.Append(new Paragraph(new Run(new Text(d.extentOfDamage))));
                tr.Append(tc2);


                TableCell tc3 = new TableCell();
                tc3.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
                tc3.Append(new Paragraph(new Run(new Text(d.inconvenienceSuffered))));
                tr.Append(tc3);


                TableCell tc4 = new TableCell();
                tc4.Append(new TableCellProperties(
                    new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
                tc4.Append(new Paragraph(new Run(new Text(d.costOfRepair))));
                tr.Append(tc4);

                // Append the table row to the table.
                table.Append(tr);
            }

            return table;
        }

        public TableRow CreateTableHeaders()
        {
            TableRow tr = new TableRow();

            TableCell tc0 = new TableCell();
            tc0.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
            tc0.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text("Item"))));
            tr.Append(tc0);


            TableCell tc1 = new TableCell();
            tc1.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
            tc1.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text("What room is it in"))));
            tr.Append(tc1);


            TableCell tc2 = new TableCell();
            tc2.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
            tc2.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text("Extent of damage"))));
            tr.Append(tc2);


            TableCell tc3 = new TableCell();
            tc3.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
            tc3.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text("Inconvenience Suffered"))));
            tr.Append(tc3);


            TableCell tc4 = new TableCell();
            tc4.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));
            tc4.Append(new Paragraph(new Run(new RunProperties(new Bold()), new Text("Est. Cost of Repair"))));
            tr.Append(tc4);

            return tr;
        }

        public Paragraph AddPageBreak()
        {
            return new Paragraph(
                new Run(new Break() { Type = BreakValues.Page }));
        }

        public Paragraph AddLandscapeParagraph()
        {
            return new Paragraph(
           new ParagraphProperties(
               new SectionProperties(
                   new PageSize() { Width = (UInt32Value)12240U, Height = (UInt32Value)15840U, Orient = PageOrientationValues.Landscape },
                   new PageMargin() { Top = 720, Right = (UInt32Value)1440U, Bottom = 360, Left = (UInt32Value)1440U, Header = (UInt32Value)450U, Footer = (UInt32Value)720U, Gutter = (UInt32Value)0U })
               ))
            { RsidParagraphAddition = "00BA2F0F", RsidParagraphProperties = "00BA2F0F", RsidRunAdditionDefault = "00BA2F0F" };
            

        }
    }
}
