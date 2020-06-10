using NPOI.SS.UserModel;
using Xsheet;

namespace XSheet.Renderers.NPOI.Formats
{
    public class NPOIFormat : IFormat
    {
        public ICellStyle CellStyle { get; set; }
    }
}
