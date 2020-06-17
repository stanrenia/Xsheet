using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xsheet.Extensions;
using Xsheet.Tests.SharedDatasets;
using XSheet.Renderers.ClosedXml;
using Xunit;

namespace Xsheet.Tests
{
    public class ClosedXmlRendererTest : IDisposable
    {
        private readonly XLWorkbook _wb;

        public ClosedXmlRenderer _renderer { get; }

        private readonly ClosedXmlFormatApplier _defaultFormatApplier;
        private readonly List<Stream> _fileStreamToClose;

        public ClosedXmlRendererTest()
        {
            _wb = new XLWorkbook();
            _defaultFormatApplier = new ClosedXmlFormatApplier();
            _renderer = new ClosedXmlRenderer(_wb, _defaultFormatApplier);
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

        [Fact]
        public void Should_Renderer_Matrix_With_ClosedXmlFormat()
        {
            // GIVEN
            const string Lastname = "Lastname";
            const string Firstname = "Firstname";
            const string Age = "Age";

            Func<IXLStyle, IXLStyle> style1 = (style) =>
            {
                style.Font.Italic = true;
                return style;
            };

            Func<IXLStyle, IXLStyle> style2 = (style) =>
            {
                style.Font.FontSize = 12;
                return style;
            };

            Func<IXLStyle, IXLStyle> style3Partial = (style) =>
            {
                style.Font.Bold = true;
                return style;
            };
            var style3 = style3Partial.Compose(style2);

            var ColorBlueIndex = XLColor.LightBlue;
            var ColorLightGreyIndex = XLColor.LightGray;

            Func<IXLStyle, IXLStyle> style4Partial = (style) =>
            {
                style.Fill.BackgroundColor = ColorLightGreyIndex;
                return style;
            };
            var style4 = style4Partial.Compose(style2);

            Func<IXLStyle, IXLStyle> style5Partial = (style) =>
            {
                style.Fill.BackgroundColor = ColorBlueIndex;
                return style;
            };
            var style5 = style5Partial.Compose(style4);

            const string Even = "EVEN";
            const string Odd = "ODD";

            var values = new List<RowValue> {
                new RowValue {
                    Key = Even,
                    ValuesByColName = new Dictionary<string, object> {
                        { Lastname, "Doe" },
                        { Firstname, "John" },
                        { Age, 30 }
                    }
                },
                new RowValue {
                    Key = Odd,
                    ValuesByColName = new Dictionary<string, object> {
                        { Lastname, "Clinton" },
                        { Firstname, "Bob" },
                        { Age, 41 }
                    }
                },
                new RowValue {
                    Key = Even,
                    ValuesByColName = new Dictionary<string, object> {
                        { Lastname, "Doa" },
                        { Firstname, "Johana" },
                        { Age, 36 }
                    }
                }
            };

            var mat = Matrix.With()
                .Cols()
                    .Col(label: Lastname)
                    .Col(label: Firstname, headerCellFormat: new ClosedXmlFormat(style1))
                    .Col(label: Age, dataType: DataTypes.Number)
                .Rows()
                    .Row(defaultCellFormat: new ClosedXmlFormat(style2))
                        .Format(Lastname, new ClosedXmlFormat(style3))
                        .Format(Age, new ClosedXmlFormat(style4))
                    .Row(key: Odd)
                        .Format(Age, new ClosedXmlFormat(style5))
                .RowValues(values)
                .Build();

            var ms = new MemoryStream();

            // WHEN
            _renderer.GenerateExcelFile(mat, ms);

            TestUtils.WriteDebugFile(mat, "closedxml_format", streamsToClose: _fileStreamToClose, formatApplier: _defaultFormatApplier);

            // THEN
            SharedAssertions.Assert_Format_With_Specific_FormatApplier(ms, new FormatDataset
            {
                ColorBlue = ColorBlueIndex.Color.ToARGB(),
                ColorLightGrey = ColorLightGreyIndex.Color.ToARGB(),
                ColsCount = mat.ColumnsDefinitions.Count()
            });
        }
    }
}
