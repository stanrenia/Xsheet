using NPOI.SS.UserModel;
using Xsheet;

namespace XSheet.Renderers.NPOI.Formats
{
    public interface IFormatApplier
    {
        void ApplyFormatToCell(IWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex, ICell cell);
        void ApplyFormatToCell(IWorkbook wb, ICell cell, IFormat format);
    }
}
