using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Worthy.DocumentBuilder.Text
{
    public class ParagraphElement : IDocumentElement
    {
        public DocumentElementType Type => DocumentElementType.Paragraph;
        public List<IDocumentElement> Elements { get; }

        public Style Style { get; set; }

        public ParagraphElement(Style style, string text) : this(style, new TextElement(text))
        {
        }

        public ParagraphElement(string text) : this(new TextElement(text))
        {
        }

        public ParagraphElement() : this(null, new IDocumentElement[0])
        {
        }

        public ParagraphElement(params IDocumentElement[] elements) : this (null, elements)
        {
        }

        public ParagraphElement(Style style, params IDocumentElement[] elements)
        {
            Style = style;
            Elements = elements.ToList();
            //Elements = elements
            //    .Select (e => {
            //        if (Style != null && e.Style !=null)
            //        {
            //            e.Style = Style;
            //        }
            //        return e;
            //    })
            //    .ToList();
        }
    }
}
