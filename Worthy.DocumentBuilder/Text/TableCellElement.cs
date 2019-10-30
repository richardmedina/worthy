using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worthy.DocumentBuilder.Text
{
    public class TableCellElement : IDocumentElement
    {
        public DocumentElementType Type => DocumentElementType.TableCell;

        public Style Style { get; set; }
        public List<IDocumentElement> Elements { get; set; }


        public TableCellElement() : this (null, new IDocumentElement[0])
        {
        }

        public TableCellElement(params IDocumentElement[] elements) : this (null, elements)
        {
        }

        public TableCellElement(Style style, params IDocumentElement [] elements)
        {
            Style = style;
            Elements = elements.ToList ();
            //Elements = elements
            //    .Select (e => {
            //        if (Style != null && e.Style != null)
            //        {
            //            e.Style = Style;
            //        }

            //        return e;
            //    })
            //    .ToList ();
        }
    }
}
