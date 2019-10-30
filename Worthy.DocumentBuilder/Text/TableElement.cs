using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worthy.DocumentBuilder.Text
{
    public class TableElement : IDocumentElement
    {
        public DocumentElementType Type => DocumentElementType.Table;

        public Style Style { get; set; }
        public List<TableRowElement> Rows { get; }

        public TableElement(IEnumerable<TableRowElement> elements) : this(elements.ToArray())
        {
        }

        public TableElement(Style style, IEnumerable<TableRowElement> elements) : this(style, elements.ToArray())
        {
        }

        public TableElement() : this (new TableRowElement[0])
        {
        }

        public TableElement(params TableRowElement[] rows) : this (null, rows)
        {
        }

        public TableElement(Style style, params TableRowElement [] rows)
        {
            Style = style;
            Rows = rows
                .ToList();
        }
    }
}
