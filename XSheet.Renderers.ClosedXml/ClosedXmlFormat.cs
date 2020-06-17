using ClosedXML.Excel;
using System;
using Xsheet;

namespace XSheet.Renderers.ClosedXml
{
    public class ClosedXmlFormat : IFormat
    {
        public Func<IXLStyle, IXLStyle> Stylize { get; set; }

        public ClosedXmlFormat()
        {

        }

        public ClosedXmlFormat(Func<IXLStyle, IXLStyle> stylize)
        {
            Stylize = stylize;
        }
    }
}
