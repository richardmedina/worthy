using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Worthy.DocumentBuilder.Text
{
    public interface ITextDocumentBuilder : IDocumentBuilder
    {
        DocumentElementType Type { get; }
    }
}
