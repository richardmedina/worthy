using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worthy.DocumentBuilder.Text
{
    public class TableRowElement : IDocumentElement
    {
        public DocumentElementType Type => DocumentElementType.TableRow;
        public List<TableCellElement> Cells { get; }
        public Style Style { get; set; }

        public TableRowElement() : this(null, new TableCellElement[0])
        {
        }

        public TableRowElement(params TableCellElement[] cells) : this (null, cells)
        {
        }

        public TableRowElement(Style defaultCellStyle, params TableCellElement [] cells)
        {
            Style = defaultCellStyle;
            Cells = cells.ToList ();
            //Cells = cells.Select(c => {
            //    if (Style != null && c.Style == null)
            //    {
            //        c.Style = Style;
            //    }
            //    return c;
            //}).ToList();
        }
    }
}
