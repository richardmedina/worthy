using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Worthy
{
    [TestClass]
    public class RawTests
    {
        [TestMethod]
        public void RawTest1 ()
        {
            var tableBuilder = new GeneratedCode.GeneratedClass();
            byte[] bytes;
            using (var stream = new MemoryStream())
            {
                using (var wordProcessingDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
                {
                    var mainDocumentPart = wordProcessingDocument.AddMainDocumentPart();
                    mainDocumentPart.Document = new Document();
                    mainDocumentPart.Document.Append(new Body (tableBuilder.GenerateTable ()));

                    mainDocumentPart.Document.Save();
                }
                bytes = stream.ToArray();
            }

            File.WriteAllBytes("rawtest1.docx", bytes);
        }
    }
}

namespace GeneratedCode
{
    using DocumentFormat.OpenXml.Wordprocessing;
    using DocumentFormat.OpenXml;


    public class GeneratedClass
    {
        // Creates an Table instance and adds its children.
        public Table GenerateTable()
        {
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableLook tableLook1 = new TableLook() { Val = "04A0", FirstRow = true, LastRow = false, FirstColumn = true, LastColumn = false, NoHorizontalBand = false, NoVerticalBand = true };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2394" };
            GridColumn gridColumn2 = new GridColumn() { Width = "2394" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2394" };
            GridColumn gridColumn4 = new GridColumn() { Width = "2394" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);
            tableGrid1.Append(gridColumn4);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "00607D74", RsidTableRowProperties = "0055020F", ParagraphId = "2D079EFF", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "1F497D", ThemeFill = ThemeColorValues.Text2 };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(shading1);
            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "6ED85602", TextId = "77777777" };

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "1F497D", ThemeFill = ThemeColorValues.Text2 };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);
            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "7C687A4B", TextId = "77777777" };

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };
            Shading shading3 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "1F497D", ThemeFill = ThemeColorValues.Text2 };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(shading3);
            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "4F287AEB", TextId = "77777777" };

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            TableCell tableCell4 = new TableCell();

            TableCellProperties tableCellProperties4 = new TableCellProperties();
            TableCellWidth tableCellWidth4 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };
            Shading shading4 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "1F497D", ThemeFill = ThemeColorValues.Text2 };

            tableCellProperties4.Append(tableCellWidth4);
            tableCellProperties4.Append(shading4);
            Paragraph paragraph4 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "5F0063F5", TextId = "77777777" };

            tableCell4.Append(tableCellProperties4);
            tableCell4.Append(paragraph4);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);
            tableRow1.Append(tableCell4);

            TableRow tableRow2 = new TableRow() { RsidTableRowAddition = "00607D74", RsidTableRowProperties = "00607D74", ParagraphId = "2558F9F0", TextId = "77777777" };

            TableCell tableCell5 = new TableCell();

            TableCellProperties tableCellProperties5 = new TableCellProperties();
            TableCellWidth tableCellWidth5 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties5.Append(tableCellWidth5);

            Paragraph paragraph5 = new Paragraph() { RsidParagraphMarkRevision = "0055020F", RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "186B732E", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            Bold bold1 = new Bold();

            paragraphMarkRunProperties1.Append(bold1);

            paragraphProperties1.Append(paragraphMarkRunProperties1);

            paragraph5.Append(paragraphProperties1);

            tableCell5.Append(tableCellProperties5);
            tableCell5.Append(paragraph5);

            TableCell tableCell6 = new TableCell();

            TableCellProperties tableCellProperties6 = new TableCellProperties();
            TableCellWidth tableCellWidth6 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties6.Append(tableCellWidth6);
            Paragraph paragraph6 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "6D2E7DAE", TextId = "77777777" };

            tableCell6.Append(tableCellProperties6);
            tableCell6.Append(paragraph6);

            TableCell tableCell7 = new TableCell();

            TableCellProperties tableCellProperties7 = new TableCellProperties();
            TableCellWidth tableCellWidth7 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties7.Append(tableCellWidth7);
            Paragraph paragraph7 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "2B98717A", TextId = "77777777" };

            tableCell7.Append(tableCellProperties7);
            tableCell7.Append(paragraph7);

            TableCell tableCell8 = new TableCell();

            TableCellProperties tableCellProperties8 = new TableCellProperties();
            TableCellWidth tableCellWidth8 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties8.Append(tableCellWidth8);
            Paragraph paragraph8 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "12461324", TextId = "77777777" };

            tableCell8.Append(tableCellProperties8);
            tableCell8.Append(paragraph8);

            tableRow2.Append(tableCell5);
            tableRow2.Append(tableCell6);
            tableRow2.Append(tableCell7);
            tableRow2.Append(tableCell8);

            TableRow tableRow3 = new TableRow() { RsidTableRowAddition = "00607D74", RsidTableRowProperties = "00607D74", ParagraphId = "56FE196D", TextId = "77777777" };

            TableCell tableCell9 = new TableCell();

            TableCellProperties tableCellProperties9 = new TableCellProperties();
            TableCellWidth tableCellWidth9 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties9.Append(tableCellWidth9);

            Paragraph paragraph9 = new Paragraph() { RsidParagraphMarkRevision = "0055020F", RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "617925F5", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            Color color1 = new Color() { Val = "FF0000" };

            paragraphMarkRunProperties2.Append(color1);

            paragraphProperties2.Append(paragraphMarkRunProperties2);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph9.Append(paragraphProperties2);
            paragraph9.Append(bookmarkStart1);
            paragraph9.Append(bookmarkEnd1);

            tableCell9.Append(tableCellProperties9);
            tableCell9.Append(paragraph9);

            TableCell tableCell10 = new TableCell();

            TableCellProperties tableCellProperties10 = new TableCellProperties();
            TableCellWidth tableCellWidth10 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties10.Append(tableCellWidth10);
            Paragraph paragraph10 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "7FB20390", TextId = "77777777" };

            tableCell10.Append(tableCellProperties10);
            tableCell10.Append(paragraph10);

            TableCell tableCell11 = new TableCell();

            TableCellProperties tableCellProperties11 = new TableCellProperties();
            TableCellWidth tableCellWidth11 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties11.Append(tableCellWidth11);
            Paragraph paragraph11 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "6B3FAA2A", TextId = "77777777" };

            tableCell11.Append(tableCellProperties11);
            tableCell11.Append(paragraph11);

            TableCell tableCell12 = new TableCell();

            TableCellProperties tableCellProperties12 = new TableCellProperties();
            TableCellWidth tableCellWidth12 = new TableCellWidth() { Width = "2394", Type = TableWidthUnitValues.Dxa };

            tableCellProperties12.Append(tableCellWidth12);
            Paragraph paragraph12 = new Paragraph() { RsidParagraphAddition = "00607D74", RsidRunAdditionDefault = "00607D74", ParagraphId = "046FFB08", TextId = "77777777" };

            tableCell12.Append(tableCellProperties12);
            tableCell12.Append(paragraph12);

            tableRow3.Append(tableCell9);
            tableRow3.Append(tableCell10);
            tableRow3.Append(tableCell11);
            tableRow3.Append(tableCell12);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            table1.Append(tableRow2);
            table1.Append(tableRow3);
            return table1;
        }

    }
}