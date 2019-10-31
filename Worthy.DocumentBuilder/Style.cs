using System;
using System.Collections.Generic;
using System.Text;

namespace Worthy.DocumentBuilder
{
    public class Style
    {
        public string FontName { get; set; }
        public int? FontSize { get; set; }
        public bool? Bold { get; set; }
        public bool? Italic { get; set; }
        public string ForegroundColor {get; set; }
        public Border Border
        {
            set => BorderTop = BorderRight = BorderBottom = BorderLeft = value;
        }

        public Border BorderTop { get; set; }
        public Border BorderRight { get; set; }
        public Border BorderBottom { get; set; }
        public Border BorderLeft { get; set; }

        public VerticalAlignment? VerticalAlignment { get; set; }
        public HorizontalAlignment? HorizontalAlignment { get; set; }

        public int? Width { get; set; }
        public uint? Height { get; set; }
        public string BackgroundColor { get; set; }

        public string ReferenceId { get; set; }
    }
}
