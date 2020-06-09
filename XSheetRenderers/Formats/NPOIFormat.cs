using NPOI.SS.UserModel;
using Xsheet;

namespace XSheet.Renderers.Formats
{
    public class NPOIFormat : IFormat
    {
        public ICellStyle CellStyle { get; set; }
    }
}
