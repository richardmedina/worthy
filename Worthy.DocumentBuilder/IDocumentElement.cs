using System;
using System.Collections.Generic;
using System.Text;

namespace Worthy.DocumentBuilder
{
    public interface IDocumentElement
    {
        DocumentElementType Type { get; }
        Style Style { get; set; }
    }
}
