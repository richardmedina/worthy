using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;

namespace Worthy.DocumentBuilder
{
    public class Border
    {
        public BorderType Type { get; set; }
        public string Color { get; set; }
        public uint Width { get; set; }
    }
}
