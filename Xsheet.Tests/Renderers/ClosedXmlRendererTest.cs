using ClosedXML.Excel;
using NFluent;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Xsheet.Tests.SharedDatasets;
using XSheet.Renderers.ClosedXml;
using Xunit;

namespace Xsheet.Tests
{
    public class ClosedXmlRendererTest : IDisposable
    {
        private readonly XLWorkbook _wb;

        public ClosedXmlRenderer _renderer { get; }

        private readonly List<Stream> _fileStreamToClose;

        public ClosedXmlRendererTest()
        {
            _wb = new XLWorkbook();
            _renderer = new ClosedXmlRenderer(_wb, null);
            _fileStreamToClose = new List<Stream>();
        }

        public void Dispose()
        {
            foreach (var stream in _fileStreamToClose)
            {
                stream.Close();
            }
        }

        [Fact]
        public void Should_Renderer_Matrix_Basic_Example()
        {
            // GIVEN
            var dataset = MatrixDatasets.Given_BasicExample();
            var mat = dataset.Matrix;

            var ms = new MemoryStream();

            // WHEN
            _renderer.GenerateExcelFile(mat, ms);
            ms.Close();

            TestUtils.WriteDebugFile(mat, "closedxml_basic_example", streamsToClose: _fileStreamToClose);

            // THEN
            SharedAssertions.Assert_Basic_Example(ms, dataset);
        }

        [Fact]
        public void Should_Render_Matrix_With_Cells_Lookup_On_A_Single_Matrix()
        {
            // GIVEN
            var mat = MatrixDatasets.Given_MatrixWithCellLookup(1, null);

            // WHEN
            var filename = TestUtils.WriteDebugFile(mat, "closedxml_lookup1", streamsToClose: _fileStreamToClose);
            _fileStreamToClose.First().Close();
            // THEN
            SharedAssertions.Assert_Matrix_Cells_Lookup_On_A_Single_Matrix(filename);
        }
    }
}
