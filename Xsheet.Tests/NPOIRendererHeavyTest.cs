using NFluent;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XSheet.Renderers.Formats;
using Xunit;

namespace Xsheet.Tests
{
    public class NPOIRendererHeavyTest : IDisposable
    {
        private IWorkbook _workbook;
        private readonly NPOIFormatApplier _defaultFormatApplier;
        private readonly List<Stream> _fileStreamToClose;

        private const string FILE_TEST_HEAVY_1 = "test_heavy1";

        public NPOIRendererHeavyTest()
        {
            _workbook = new XSSFWorkbook();
            _defaultFormatApplier = new NPOIFormatApplier();
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
        // Concat horizontally 100 Matrix
        public void Should_Render_Matrix_With_Heavy_Load1()
        {
            // GIVEN
            var finalTotalStyle = _workbook.CreateCellStyle();
            finalTotalStyle.FillPattern = FillPattern.SolidForeground;
            finalTotalStyle.FillForegroundColor = IndexedColors.LightOrange.Index;

            var finalMatrix = Enumerable.Range(1, 100)
                .Select(i => MatrixCellLookup(i, finalTotalStyle))
                .Aggregate(MatrixCellLookup(0, finalTotalStyle), (acc, next) =>
                {
                    return acc.ConcatX(next);
                });

            // WHEN
            var filename = TestUtils.WriteDebugFile(_defaultFormatApplier, finalMatrix, FILE_TEST_HEAVY_1, _workbook, _fileStreamToClose);

            // THEN
            var readWb = new XSSFWorkbook(File.OpenRead(filename));
            var readSheet = readWb.GetSheetAt(0);

            var row1 = readSheet.GetRow(1);
            Check.That(row1.Cells[0].StringCellValue).IsEqualTo("Mario");
            var last = 605;
            Check.That(row1.Cells[last].CellFormula).IsEqualTo("AVERAGE(10,20,30)");
            Check.That(readSheet.GetRow(2).Cells[last].CellFormula).IsEqualTo("AVERAGE(12,23,34)");
            Check.That(readSheet.GetRow(3).Cells[last].CellFormula).IsEqualTo("AVERAGE(WH2:WH3)");
        }

        static Matrix MatrixCellLookup(int index, ICellStyle finalTotalStyle)
        {
            const string Playername = "Playername";
            const string Score1 = "Score1";
            const string Score2 = "Score2";
            const string Score3 = "Score3";
            const string Total = "Total";
            const string Mean = "Mean";
            const string FinalTotal = "FinalTotal";

            return Matrix.With().Key(index: index)
                            .Cols(new List<ColumnDefinition> {
                    new ColumnDefinition { Name = Playername, Label = "Player name" },
                    new ColumnDefinition { Name = Score1, Label = "Score 1", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = Score2, Label = "Score 2", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = Score3, Label = "Score 3", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = Total, Label = "Total", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = Mean, Label = "Mean", DataType = DataTypes.Number },
                            })
                            .Rows(new List<RowDefinition> {
                    new RowDefinition {
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>> {
                            { Total, (mat, cell) => {
                                var row = mat.Row(cell);
                                return Convert.ToDouble(row.Col(Score1).Value)
                                + Convert.ToDouble(row.Col(Score2).Value)
                                + Convert.ToDouble(row.Col(Score3).Value);
                            }},
                            { Mean, (mat, cell) => {
                                var row = mat.Row(cell);
                                return $"=AVERAGE({row.Col(Score1).Value},{row.Col(Score2).Value},{row.Col(Score3).Value})";
                            } },
                        }
                    },
                    new RowDefinition {
                        Key = FinalTotal,
                        DefaultCellFormat = new NPOIFormat { CellStyle = finalTotalStyle },
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>
                        {
                            { Playername, (mat, cell) => "TOTAL" },
                            { Score1, (mat, cell) => mat.Col(cell).Values.Select(v => Convert.ToDouble(v)).Sum() },
                            { Score2, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.Row(cell, cell.RowIndex - 1).Col(Score2).Address})" },
                            { Score3, (mat, cell) => {
                                var adresses = mat.Col(cell).Cells.SkipLast(1).Select(c => c.Address);
                                var formula = string.Join('+', adresses);
                                return $"={formula}";
                            }},
                            { Total, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.Row(cell, cell.RowIndex - 1).Col(Total).Address})" },
                            { Mean, (mat, cell) => $"=AVERAGE({mat.Col(cell).Cells[0].Address}:{mat.Row(cell, cell.RowIndex - 1).Col(Mean).Address})" },
                        }
                    },
                            })
                            .RowValues(new List<RowValue>
                            {
                    new RowValue { ValuesByColName = new Dictionary<string, object> {
                        { Playername, "Mario" }, { Score1, 10 }, { Score2, 20 }, { Score3, 30 }
                    } },
                    new RowValue { ValuesByColName = new Dictionary<string, object> {
                        { Playername, "Luigi" }, { Score1, 12 }, { Score2, 23 }, { Score3, 34 }
                    } },
                    new RowValue { Key = FinalTotal }
                            })
                            .Build();
        }
    }
}
