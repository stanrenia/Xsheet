using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using Xsheet;

namespace XSheetRenderers
{
    public class MatrixRenderer
    {
        private readonly NPOIRenderer _renderer;

        public MatrixRenderer(NPOIRenderer renderer)
        {
            _renderer = renderer;
        }

        public ISheet GenerateExcelWorksheet(Matrix mat)
        {
            return _renderer.GenerateExcelWorksheet(mat);
        }
    }
}
