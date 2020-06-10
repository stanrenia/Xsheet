using NPOI.SS.UserModel;

namespace XSheet.Renderers.NPOI.Formats
{
    public class NPOIFormatApplier : BaseFormatApplier<NPOIFormat>
    {
        public override void ApplyFormatToCell(IWorkbook wb, ICell cell, NPOIFormat format)
        {
            cell.CellStyle = format.CellStyle;
        }
    }
}
