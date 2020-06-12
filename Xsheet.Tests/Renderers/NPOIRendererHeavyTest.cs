using NFluent;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XSheet.Renderers.NPOI.Formats;
using Xunit;

namespace Xsheet.Tests
{
    public class NPOIRendererHeavyTest : IDisposable
    {
        private IWorkbook _workbook;
        private readonly NPOIFormatApplier _defaultFormatApplier;
        private readonly List<Stream> _fileStreamToClose;

        private const string FILE_TEST_HEAVY_1 = "test_heavy1";
        private const string FILE_TEST_HEAVY_2 = "test_heavy2";

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

        [Fact]
        public void Should_Render_Weekly_Game_Score_Report()
        {
            // GIVEN
            IFont BoldFont()
            {
                var boldFont = _workbook.CreateFont();
                boldFont.IsBold = true;
                return boldFont;
            }

            var monthStyle = _workbook.CreateCellStyle();
            monthStyle.FillPattern = FillPattern.SolidForeground;
            monthStyle.FillForegroundColor = IndexedColors.LightBlue.Index;
            monthStyle.SetFont(BoldFont());

            var totalStyle = _workbook.CreateCellStyle();
            totalStyle.FillPattern = FillPattern.SolidForeground;
            totalStyle.FillForegroundColor = IndexedColors.LightOrange.Index;
            totalStyle.SetFont(BoldFont());

            var valueMapVar1 = GetVariation("Score1", "Score3");
            var valueMapVar2 = GetVariation("Score2", "Score3");
            var valueMapMonthlyScore = ValueMapMonthlyScore();
            var valueMapTotalScore = ValueMapTotalScore();
            var valueMapAverageScore = ValueMapAverageScore();

            var rowValues = GetScoreValues()
            .Select(score => new RowValue {
                Key = score.Month != null ? "MONTH" : "WEEK",
                ValuesByColName = score.ToDictionary()
            }).Concat(new List<RowValue> {
                new RowValue { Key = "TOTAL" },
                new RowValue { Key = "AVERAGE" },
            })
            .ToList();

            var mat = Matrix.With().Key(index: 0)
                .Cols(new List<ColumnDefinition> {
                    new ColumnDefinition { Name = "Week", Label = "Week", DataType = DataTypes.Text },
                    new ColumnDefinition { Name = "Score1", Label = "Score 1", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = "Var1", Label = "Score1/Score3", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = "Score2", Label = "Score 2", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = "Var2", Label = "Score2/Score3", DataType = DataTypes.Number },
                    new ColumnDefinition { Name = "Score3", Label = "Score 3", DataType = DataTypes.Number }
                })
                .Rows(new List<RowDefinition>
                {
                    new RowDefinition {
                        Key = "WEEK",
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>
                        {
                            { "Var1", valueMapVar1 },
                            { "Var2", valueMapVar2 }
                        }
                    },
                    new RowDefinition {
                        Key = "MONTH",
                        DefaultCellFormat = new NPOIFormat { CellStyle = monthStyle },
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>
                        {
                            { "Var1", valueMapVar1 },
                            { "Var2", valueMapVar2 },
                            { "Score1", valueMapMonthlyScore },
                            { "Score2", valueMapMonthlyScore },
                            { "Score3", valueMapMonthlyScore }
                        }
                    },
                    new RowDefinition
                    {
                        Key = "TOTAL",
                        DefaultCellFormat = new NPOIFormat { CellStyle = totalStyle },
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>
                        {
                            { "Week", (mat, cell) => "TOTAL" },
                            { "Var1", valueMapVar1 },
                            { "Var2", valueMapVar2 },
                            { "Score1", valueMapTotalScore },
                            { "Score2", valueMapTotalScore },
                            { "Score3", valueMapTotalScore },
                        }
                    },
                    new RowDefinition
                    {
                        Key = "AVERAGE",
                        DefaultCellFormat = new NPOIFormat { CellStyle = totalStyle },
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>
                        {
                            { "Week", (mat, cell) => "AVERAGE" },
                            { "Var1", valueMapVar1 },
                            { "Var2", valueMapVar2 },
                            { "Score1", valueMapAverageScore },
                            { "Score2", valueMapAverageScore },
                            { "Score3", valueMapAverageScore },
                        }
                    }
                })
                .RowValues(rowValues)
                .Build();

            // WHEN
            var filename = TestUtils.WriteDebugFile(_defaultFormatApplier, mat, FILE_TEST_HEAVY_2, _workbook, _fileStreamToClose);

            // THEN

        }

        private List<ScoreValue> GetScoreValues()
        {
            var random = new Random();
            var monthByWeek = new List<int> { 4, 8, 12, 17, 22, 26, 30, 35, 39, 43, 47, 52 };
            return Enumerable.Range(1, 52).SelectMany(i =>
            {
                var scores = new List<ScoreValue>
                {
                    new ScoreValue
                    {
                        Week = i,
                        Score1 = random.Next(999, 99999),
                        Score2 = random.Next(999, 99999),
                        Score3 = random.Next(999, 99999),
                    }
                };

                if (monthByWeek.Contains(i))
                {
                    var monthName = new DateTime(2010, monthByWeek.IndexOf(i)+1, 1).ToString("MMMM");
                    scores.Add(new ScoreValue { Month = monthName });
                }

                return scores;
            }).ToList();
        }

        private static Func<Matrix, MatrixCellValue, object> ValueMapAverageScore()
        {
            return (mat, cell) => $"=AVERAGE({mat.Col(cell).CellsOfRowKey("WEEK").Addresses()})";
        }

        private static Func<Matrix, MatrixCellValue, object> ValueMapTotalScore()
        {
            return (mat, cell) => $"=SUM({mat.Col(cell).CellsOfRowKey("WEEK").Addresses()})";
        }

        private static Func<Matrix, MatrixCellValue, object> ValueMapMonthlyScore()
        {
            return (mat, cell) =>
            {
                var addresses = mat.Col(cell)
                    .CellsBetween(cell, mat.RowIndexOfPrevious("MONTH", cell))
                    .Addresses();
                return $"=SUM({addresses})";
            };
        }

        private static Func<Matrix, MatrixCellValue, object> GetVariation(string fromColName, string toColName)
        {
            return (mat, cell) => {
                var from = mat.Row(cell).Col(fromColName).Address;
                var to = mat.Row(cell).Col(toColName).Address;
                return $"=(({to}-{from})/{from})*100";
            };
        }
    }

    internal class ScoreValue
    {
        public int Week { get; internal set; }
        public int Score1 { get; internal set; }
        public int Score2 { get; internal set; }
        public int Score3 { get; internal set; }
        public string Month { get; internal set; }

        public Dictionary<string, object> ToDictionary()
        {
            return new Dictionary<string, object>
            {
                { "Week", Month is null ? Week.ToString() : Month },
                { "Score1", Score1 },
                { "Score2", Score2 },
                { "Score3", Score3 }
            };
        }
    }
}
