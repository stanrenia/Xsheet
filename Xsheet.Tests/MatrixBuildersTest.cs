using NFluent;
using System.Collections.Generic;
using System.Linq;
using Xsheet.Formats;
using Xunit;

namespace Xsheet.Tests
{
    public class MatrixBuildersTest
    {
        public MatrixBuildersTest()
        {

        }


        [Fact]
        public void Should_Build_Matrix_With_Fluent_Api()
        {
            // GIVEN
            const string colA = "colA";
            const string colB = "colB";
            const string row1 = "ROW_1";
            const string row2 = "ROW_2";
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { colA, 11 }, { colB, 44 } } },
                new RowValue { Key = row1, ValuesByColName = new Dictionary<string, object>{ { colA, 22 }, { colB, 55 } } },
                new RowValue { Key = row2, ValuesByColName = new Dictionary<string, object>{ { colA, 33 }, { colB, 66 } } },
            };

            IFormat format1 = new BasicFormat { FontSize = 12 };
            IFormat format2 = new BasicFormat { FontSize = 11 };
            IFormat format3 = new BasicFormat { FontSize = 10 };
            IFormat format4 = new BasicFormat { FontSize = 9 };

            // WHEN
            var mat = Matrix.Builder.New
                .Cols()
                    .Col(name: colA, label: "I'm A", dataType: DataTypes.Boolean, headerCellFormat: format1)
                    .Col(colB, "I'm B")
                .Rows()
                    .Row(defaultCellFormat: format2)
                        .Format(colA, format3)
                        .ValueMap(colB, (mat, cell) => "Default B")
                    .Row(key: row1, defaultCellFormat: format3)
                        .ValueMap(colA, (mat, cell) => "Row1 A")
                        .Format(colB, format4)
                        .ValueMap(colB, (mat, cell) => "Row1 B")
                    .Row(row2, format4)
                        .ValueMap(colA, (mat, cell) => "Row2 A")
                        .ValueMap(colB, (mat, cell) => "Row2 B")
                .RowValues(values)
                .Build();

            // THEN
            Check.That(mat.ColumnsDefinitions).HasSize(2);
            Check.That(mat.RowsDefinitions).HasSize(3);
            Check.That(mat.RowValues).HasSize(3);

            // Check Columns
            var c1 = mat.ColumnsDefinitions.ElementAt(0);
            Check.That(c1.Name).IsEqualTo(colA);
            Check.That(c1.Label).IsEqualTo("I'm A");
            Check.That(c1.DataType).IsEqualTo(DataTypes.Boolean);
            Check.That(c1.HeaderCellFormat).IsEqualTo(format1);
            var c2 = mat.ColumnsDefinitions.ElementAt(1);
            Check.That(c2.Name).IsEqualTo(colB);
            Check.That(c2.Label).IsEqualTo("I'm B");

            // Check Rows
            var r = mat.RowsDefinitions.ElementAt(0);
            Check.That(r.DefaultCellFormat).IsEqualTo(format2);
            Check.That(r.FormatsByColIndex[0]).IsEqualTo(format3);
            Check.That(r.ValuesMapping[colB].Invoke(null, null)).IsEqualTo("Default B");
            var r1 = mat.RowsDefinitions.ElementAt(1);
            Check.That(r1.Key).IsEqualTo(row1);
            Check.That(r1.DefaultCellFormat).IsEqualTo(format3);
            Check.That(r1.ValuesMapping[colA].Invoke(null, null)).IsEqualTo("Row1 A");
            Check.That(r1.ValuesMapping[colB].Invoke(null, null)).IsEqualTo("Row1 B");
            Check.That(r1.FormatsByColIndex[1]).IsEqualTo(format4);
            var r2 = mat.RowsDefinitions.ElementAt(2);
            Check.That(r2.Key).IsEqualTo(row2);
            Check.That(r2.DefaultCellFormat).IsEqualTo(format4);
            Check.That(r2.ValuesMapping[colA].Invoke(null, null)).IsEqualTo("Row2 A");
            Check.That(r2.ValuesMapping[colB].Invoke(null, null)).IsEqualTo("Row2 B");

            // Check Values
            List<MatrixCellValue> cells = mat.RowValues.SelectMany(rv => rv.Cells).ToList();
            Check.That(cells).HasSize(6);
            Check.That(cells.Extracting(nameof(MatrixCellValue.Value))).ContainsExactly(11, 44, 22, 55, 33, 66);
            Check.That(cells.Extracting(nameof(MatrixCellValue.RowIndex))).ContainsExactly(1, 1, 2, 2, 3, 3);
            Check.That(cells.Extracting(nameof(MatrixCellValue.ColIndex))).ContainsExactly(0, 1, 0, 1, 0, 1);
            Check.That(cells.Extracting(nameof(MatrixCellValue.Address))).ContainsExactly("A2", "B2", "A3", "B3", "A4", "B4");
        }
    }
}
