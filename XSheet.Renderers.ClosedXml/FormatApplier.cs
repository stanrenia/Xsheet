using ClosedXML.Excel;
using Xsheet;

namespace XSheet.Renderers.ClosedXml
{
    public interface IFormatApplier
    {
        void ApplyFormatToCell(IXLWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex, IXLCell cell);
        void ApplyFormatToCell(IXLWorkbook wb, IXLCell cell, IFormat format);
    }
}
