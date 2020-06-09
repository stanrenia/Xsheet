using System;
using System.Collections.Generic;
using System.Linq;

namespace Xsheet.Tests.SharedDatasets
{
    public static class MatrixDatasets
    {
        public static MatrixDataset Given_BasicExample()
        {
            var cols = new List<ColumnDefinition>
            {
                new ColumnDefinition { Label = "Lastname", DataType = DataTypes.Text },
                new ColumnDefinition { Label = "Firstname", DataType = DataTypes.Text },
            };

            var values = new List<RowValue> {
                new RowValue {
                    ValuesByColName = new Dictionary<string, object> {
                        { "Lastname", "Doe" },
                        { "Firstname", "John" }
                    }
                }
            };

            var mat = Matrix.With()
                .Dimensions(2, 2)
                .Cols(cols)
                .RowValues(values)
                .Build();

            return new MatrixDataset
            {
                Cols = cols,
                Values = values,
                Matrix = mat
            };
        }

        public static Matrix Given_MatrixWithCellLookup(int index, IFormat format)
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
                        DefaultCellFormat = format,
                        ValuesMapping = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>
                        {
                            { Playername, (mat, cell) => "TOTAL" },
                            { Score1, (mat, cell) => mat.Col(cell).Values.Select(v => Convert.ToDouble(v)).Sum() },
                            { Score2, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Score2).Address})" },
                            { Score3, (mat, cell) => {
                                return $"={mat.Col(cell).Cells.SkipLast(1).Addresses("+")}";
                            }},
                            { Total, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Total).Address})" },
                            { Mean, (mat, cell) => $"=AVERAGE({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Mean).Address})" },
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
