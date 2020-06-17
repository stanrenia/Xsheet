using NPOI.SS.UserModel;
using Xsheet;

namespace XSheet.Renderers.NPOI.Formats
{
    public class NPOIFormat : IFormat
    {
        public NPOIFormat()
        {

        }

        public NPOIFormat(ICellStyle cellStyle)
        {
            CellStyle = cellStyle;
        }

        public ICellStyle CellStyle { get; set; }
    }
}
