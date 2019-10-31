using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;

namespace Worthy.DocumentBuilder
{
    public interface IDocumentBuilder
    {
        List<IDocumentElement> Elements { get; }
        byte [] Build();
    }
}
