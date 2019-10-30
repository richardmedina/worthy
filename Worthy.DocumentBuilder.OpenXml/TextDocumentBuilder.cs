using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Worthy.DocumentBuilder;
using Worthy.DocumentBuilder.Text;
using Border = Worthy.DocumentBuilder.Border;
using Style = Worthy.DocumentBuilder.Style;
using WordText = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Worthy.DocumentBuilder.OpenXml
{
    public class TextDocumentBuilder : ITextDocumentBuilder
    {
        public List<IDocumentElement> Elements { get; }
        public DocumentElementType Type => DocumentElementType.Document;

        private Body body;

        private IDictionary<DocumentElementType, string> exceptionConditions = new Dictionary<DocumentElementType, string>
        {
            { DocumentElementType.Text, "Could not render isolated Text Element" }
        };

        public TextDocumentBuilder() : this(new IDocumentElement[0])
        {
        }
        public TextDocumentBuilder(params IDocumentElement[] elements)
        {
            Elements = elements.ToList ();
        }

        public byte [] RenderToByteArray()
        {
            var body = BuildBody();

            byte[] bytes;
            using (var stream = new MemoryStream())
            {
                using (var wordProcessingDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
                {
                    var mainDocumentPart = wordProcessingDocument.AddMainDocumentPart();
                    mainDocumentPart.Document = new Document();
                    mainDocumentPart.Document.Append(body);

                    mainDocumentPart.Document.Save();
                }
                bytes = stream.ToArray();
            }
            return bytes;
        }

        #region Private Members
        private Body BuildBody()
        {
            body = new Body();

            foreach (var element in Elements)
            {
                if (exceptionConditions.ContainsKey(element.Type))
                    throw new ArgumentException(exceptionConditions[element.Type]);

                switch (element.Type)
                {
                    case DocumentElementType.Paragraph:
                        body.Append(GetParagraph(element as ParagraphElement));
                        break;
                    case DocumentElementType.Table:
                        body.Append(GetTable(element as TableElement));
                        break;
                }
            }

            return body;
        }

        Paragraph GetParagraph(ParagraphElement paragraph)
        {
            var para = new Paragraph();
            var run = new Run();

            foreach (TextElement text in paragraph.Elements)
            {
                var style = text.Style;

                if (style == null)
                {
                    style = paragraph.Style;
                }


                if (style != null)
                {
                    var props = new RunProperties();
                    if (!string.IsNullOrWhiteSpace (style.ForegroundColor))
                    {
                        props.Append(new Color()
                        {
                            Val = style.ForegroundColor
                        });
                    }
                    run.Append(props);

                    if (!string.IsNullOrWhiteSpace (style.FontName))
                    {
                        props.Append(new RunFonts
                        {
                            Ascii = style.FontName
                        });
                    }

                    if (style.FontSize.HasValue)
                    {
                        props.Append(new FontSize
                        {
                            Val = $"{style.FontSize.Value * 2}"
                        });
                    }

                    if (style.Bold == true)
                    {
                        props.Append(new Bold ());
                    }

                    if (style.Italic == true)
                    {
                        props.Append(new Italic());
                    }
                }

                run.Append(GetConcatenatedText(text));
            }

            para.Append(run);

            return para;
        }

        TBorderType GetBorder<TBorderType> (Border border) where TBorderType : DocumentFormat.OpenXml.Wordprocessing.BorderType
        {
            if (border == null) return null;

            TBorderType borderType = null;

            borderType = Activator.CreateInstance<TBorderType>();

            borderType.Val = new EnumValue<BorderValues>((BorderValues)border.Type);
            borderType.Size = border.Width;
            borderType.Color = border.Color;

            return borderType;
        }

        Table GetTable(TableElement tableElement)
        {
            var table = new Table();

            if (tableElement.Style != null)
            {
                var style = tableElement.Style;

                var borders = new OpenXmlLeafElement[] {
                    GetBorder<LeftBorder>(style.BorderLeft),
                    GetBorder<TopBorder>(style.BorderTop),
                    GetBorder<RightBorder>(style.BorderRight),
                    GetBorder<BottomBorder>(style.BorderBottom)
                }
                .Where(b => b != null);

                var tableBorders = new TableBorders(borders);

                var tableProperties = new TableProperties(tableBorders);
                table.Append(tableProperties);
            }

            var rows = tableElement.Rows.Select(row => {
                var tableRow = new TableRow(row.Cells.Select(cell =>
                {
                    var tableCell = new TableCell(cell.Elements.Select(e =>
                    {
                        e.Style = e.Style ?? (cell.Style ?? row.Style);

                        return GetParagraph(e as ParagraphElement);
                    }));

                    var style = cell.Style ?? row.Style;

                    if (style != null)
                    {
                        var props = new TableCellProperties();

                        var borders = new OpenXmlLeafElement[] {
                            GetBorder<LeftBorder>(style.BorderLeft),
                            GetBorder<TopBorder>(style.BorderTop),
                            GetBorder<RightBorder>(style.BorderRight),
                            GetBorder<BottomBorder>(style.BorderBottom)
                        }
                        .Where(b => b != null);

                        if (borders.Any())
                        {
                            props.Append(new TableBorders(borders));
                        }

                        if (style.BackgroundColor != null)
                        {
                            props.Append(new Shading
                            {
                                Color = "auto",
                                Fill = style.BackgroundColor,
                                Val = ShadingPatternValues.Clear
                            });                            
                        }

                        if (style.VerticalAlignment.HasValue)
                        {
                            props.Append(new TableCellVerticalAlignment
                            {
                                //Val = (TableVerticalAlignmentValues) style.VerticalAlignment.Value
                                Val = new EnumValue<TableVerticalAlignmentValues>((TableVerticalAlignmentValues) style.VerticalAlignment.Value)
                            });
                        }

                        if (style.HorizontalAlignment.HasValue)
                        {
                            props.Append(new Justification {
                                //Val = (JustificationValues)style.HorizontalAlignment
                                Val = new EnumValue<JustificationValues>((JustificationValues)style.HorizontalAlignment.Value)
                            });
                        }

                        tableCell.Append(props);
                    }

                    //if (cell.Style != null)
                    //{
                    //    if (row.Style.BackgroundColor != null)
                    //    {
                    //        var props = new TableCellProperties();
                    //        props.Append(new Shading
                    //        {
                    //            Color = "auto",
                    //            Fill = row.Style.BackgroundColor,
                    //            Val = ShadingPatternValues.Clear
                    //        });
                    //        tableCell.Append(props);
                    //    }
                    //}

                    return tableCell;
                }));

                //if (row.Style != null)
                //{
                //    var borders = new OpenXmlLeafElement[] {
                //            GetBorder<LeftBorder>(row.Style.BorderLeft),
                //            GetBorder<TopBorder>(row.Style.BorderTop),
                //            GetBorder<RightBorder>(row.Style.BorderRight),
                //            GetBorder<BottomBorder>(row.Style.BorderBottom)
                //        }
                //            .Where(b => b != null);

                //    if (borders.Any())
                //    {
                //        var props = new TableRowProperties();

                //        props.Append(new TableBorders(borders));
                //        tableRow.Append(props);
                //    }
                //}

                return tableRow;
            });

            table.Append(rows);

            return table;
        }

        OpenXmlElement GetConcatenatedText(TextElement textElement)
        {
            var texts = new List<OpenXmlLeafElement>();
            var chunks = textElement.Text.Split('\n');

            for (var i = 0; i < chunks.Length; i++)
            {
                texts.Add(new WordText(chunks[i]));
                if (i < chunks.Length - 1)
                {
                    texts.Add(new Break());
                }
            }

            return new Run (texts);
        }

        #endregion

    }
}
