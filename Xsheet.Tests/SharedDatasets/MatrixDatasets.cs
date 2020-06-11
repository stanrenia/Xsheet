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

            return Matrix.With()
                .Key(index: index)
                .Cols()
                    .Col(Playername, "Player name")
                    .Col(Score1, "Score 1", DataTypes.Number)
                    .Col(Score2, "Score 2", DataTypes.Number)
                    .Col(Score3, "Score 3", DataTypes.Number)
                    .Col(Total, "Total", DataTypes.Number)
                    .Col(Mean, "Mean", DataTypes.Number)
                .Rows()
                    .Row()
                        .ValueMap(Total, (mat, cell) =>
                        {
                            var row = mat.Row(cell);
                            return Convert.ToDouble(row.Col(Score1).Value)
                            + Convert.ToDouble(row.Col(Score2).Value)
                            + Convert.ToDouble(row.Col(Score3).Value);
                        })
                        .ValueMap(Mean, (mat, cell) =>
                        {
                            var row = mat.Row(cell);
                            return $"=AVERAGE({row.Col(Score1).Value},{row.Col(Score2).Value},{row.Col(Score3).Value})";
                        })
                    .Row(FinalTotal, format)
                        .ValueMap(Playername, (mat, cell) => "TOTAL")
                        .ValueMap(Score1, (mat, cell) => mat.Col(cell).Values.Select(v => Convert.ToDouble(v)).Sum())
                        .ValueMap(Score2, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Score2).Address})")
                        .ValueMap(Score3, (mat, cell) => $"={mat.Col(cell).Cells.SkipLast(1).Addresses("+")}")
                        .ValueMap(Total, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Total).Address})")
                        .ValueMap(Mean, (mat, cell) => $"=AVERAGE({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Mean).Address})")
                .RowValues(new List<RowValue> {
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
