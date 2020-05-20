using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using Xsheet;

namespace XSheetRenderers
{
    public class NPOIRenderer
    {
        private readonly IWorkbook _wb;

        public NPOIRenderer(IWorkbook wb)
        {
            _wb = wb;
        }

        public ISheet GenerateExcelWorksheet(Matrix mat)
        {
            return null;
        }
    }
}
