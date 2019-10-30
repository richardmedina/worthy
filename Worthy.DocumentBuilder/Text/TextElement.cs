using System;
using System.Collections.Generic;
using System.Text;

namespace Worthy.DocumentBuilder.Text
{
    public class TextElement: IDocumentElement
    {
        public DocumentElementType Type => DocumentElementType.Text;
        public string Text { get; set; }
        public Style Style { get; set; }

        public TextElement() : this (null, string.Empty)
        {
        }

        public TextElement(string text) : this (null, text)
        {
        }

        public TextElement(Style style, string text)
        {
            Text = text;
            Style = style;
        }
    }
}
