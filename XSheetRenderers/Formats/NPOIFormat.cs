using NPOI.SS.UserModel;
using Xsheet;

namespace XSheet.Renderers.Formats
{
    public class NPOIFormat : Format
    {
        public ICellStyle CellStyle { get; set; }
    }
}
