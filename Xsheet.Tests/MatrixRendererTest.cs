using NFluent;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using XSheetRenderers;
using Xunit;

namespace Xsheet.Tests
{
    public class MatrixRendererTest
    {
        private readonly IWorkbook wb;
        private readonly NPOIRenderer npoiRenderer;
        private readonly MatrixRenderer renderer;

        public MatrixRendererTest()
        {
            wb = new XSSFWorkbook();
            npoiRenderer = new NPOIRenderer(wb);
            renderer = new MatrixRenderer(npoiRenderer);
        }

        [Fact]
        public void Should_Renderer_Matrix_Of_Same_Size()
        {
            // GIVEN
            var mat = new Matrix();

            // WHEN
            var sheet = renderer.GenerateExcelWorksheet(mat);

            // THEN
            Check.That(sheet).IsNotNull();
            Check.That(sheet.FirstRowNum).IsEqualTo(0);
            Check.That(sheet.LastRowNum).IsEqualTo(5);
            Check.That(sheet.LeftCol).IsEqualTo(0);

        }
    }
}
