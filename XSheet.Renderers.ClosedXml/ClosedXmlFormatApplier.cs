using ClosedXML.Excel;
using System;
using System.Linq;
using Xsheet;
using Xsheet.Extensions;
using Xsheet.Formats;

namespace XSheet.Renderers.ClosedXml
{
    public class ClosedXmlFormatApplier : IFormatApplier
    {
        public void ApplyFormatToCell(IXLWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex, IXLCell cell)
        {
            var formats = FormatMerger.GetFormatsOrderedByLessPrioritary<ClosedXmlFormat>(defaultRowDef, rowDef, columnIndex);
            formats.Reverse();
            Func <IXLStyle, IXLStyle> f = (s) => null;
            var composedStylize = formats.Aggregate(f, (mergedFormat, nextFormat) =>
            {
                return mergedFormat.Compose(nextFormat.Stylize);
            });

            ApplyFormatToCell(wb, cell, new ClosedXmlFormat(composedStylize));
        }

        public void ApplyFormatToCell(IXLWorkbook wb, IXLCell cell, IFormat format)
        {
            if (format is ClosedXmlFormat closedXmlFormat)
            {
                closedXmlFormat.Stylize.Invoke(cell.Style);
            }
        }
    }
}
