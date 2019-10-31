using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Worthy.DocumentBuilder;
using Worthy.DocumentBuilder.OpenXml;
using Worthy.DocumentBuilder.Text;

namespace Worthy
{
    [TestClass]
    public class WordGenerationTests
    {

        private static string FontName = "Noto Sans";
        private static int FontSize = 12;

        [TestMethod]
        public void CreateDocumentBuilder()
        {
            IDocumentBuilder docBuilder = new WordDocumentBuilder();

            var para = new ParagraphElement(
                new Style {
                    ForegroundColor = "green"
                },
                new TextElement("Hello World, this is the paragraph content. "),
                new TextElement(new Style { ForegroundColor = "#ff0000" }, "This is another text next to the previous one")
            );

            //var para2 = new ParagraphElement(
            //    new TextElement("Hello World, this is the paragraph content. "),
            //    new TextElement("This is another text next to the previous one")
            //);

            var rows = new List<TableRowElement> {
                new TableRowElement(
                    new Style {
                        ForegroundColor = "white",
                        BackgroundColor = "2e2e38",
                        FontName = FontName,
                        FontSize = FontSize,
                        Bold = true
                    },
                    new TableCellElement(new ParagraphElement("Hello")),
                    new TableCellElement(new ParagraphElement("World")),
                    new TableCellElement(new ParagraphElement("How")),
                    new TableCellElement(new ParagraphElement("Are")),
                    new TableCellElement(new ParagraphElement("You")),
                    new TableCellElement(new ParagraphElement("Doing"))
                )
            };

            rows.AddRange(Enumerable.Range(1, 10)
                .Select(n => new TableRowElement(
                    new Style
                    {
                        Border = new Border
                        {
                            Color = "444",
                            Type = BorderType.Thick,
                            Width = 10
                        },
                        VerticalAlignment = VerticalAlignment.Center,
                        BackgroundColor = n % 2 == 0 ? "#f6f6f6" : null,
                        FontName = FontName,
                        FontSize = FontSize,
                        Height = 70
                    },
                    new TableCellElement(new ParagraphElement("Hello")),
                    new TableCellElement(new ParagraphElement("World")),
                    new TableCellElement(new ParagraphElement("How")),
                    new TableCellElement(new ParagraphElement("Are")),
                    new TableCellElement(new ParagraphElement("You")),
                    new TableCellElement(new ParagraphElement("Doing"))
                )));

            var table = new TableElement(
                new Style
                {
                    Border = new Border
                    {
                        Type = BorderType.Thick,
                        Width = 10,
                        Color = "333333"
                    },
                    Width = 50
                },
                rows
            );

            docBuilder.Elements.Add(para);
            //docBuilder.Elements.Add(para2);
            docBuilder.Elements.Add(table);


            var bytes = docBuilder.RenderToByteArray();

            File.WriteAllBytes("myfile.docx", bytes);
        }
    }
}
