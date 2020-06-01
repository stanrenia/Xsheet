using NFluent;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace Xsheet.Tests
{
    public class MatrixTest
    {
        public MatrixTest()
        {

        }

        [Theory]
        [InlineData(1, 1, true)]
        [InlineData(2, 2, true)]
        [InlineData(0, 0, false)]
        [InlineData(0, 2, false)]
        [InlineData(-1, 2, false)]
        [InlineData(2, 0, false)]
        [InlineData(2, -1, false)]
        public void Should_Build_Matrix_With_Dimension_Greater_Than_Zero(int x, int y, bool expected)
        {
            bool succeeded = true;
            try
            {
                var mat = Matrix.With()
                    .Dimensions(x, y)
                    .Build();
            }
            catch
            {
                succeeded = false;
            }

            Check.That(succeeded).IsEqualTo(expected);
        }

        [Fact]
        public void Should_Build_Matrix_With_Columns_Definition_And_Rows_Count()
        {
            // GIVEN
            var cols = new List<ColumnDefinition>
            {
                new ColumnDefinition { Name = "ColA" }
            };

            // WHEN
            var mat = Matrix.With()
                .RowsCount(10)
                .Cols(cols)
                .Build();

            // THEN
            Check.That(mat.ColumnsDefinitions).IsEquivalentTo(cols);
            Check.That(mat.CountOfRows).IsEqualTo(10);
            Check.That(mat.CountOfColumns).IsEqualTo(1);
        }

        [Fact]
        public void Should_Throw_When_Build_Matrix_With_Columns_Definition_Without_Name_And_Label()
        {
            // GIVEN
            var cols = new List<ColumnDefinition>
            {
                new ColumnDefinition { }
            };

            var matBuilder = Matrix.With()
                .RowsCount(10)
                .Cols(cols);

            // WHEN
            Action action = () => matBuilder.Build();

            // THEN
            Check.ThatCode(action).Throws<Exception>();
        }

        [Fact]
        public void Should_Build_Matrix_With_Rows_Definition_And_Cols_Count()
        {
            // GIVEN
            var rows = new List<RowDefinition>
            {
                new RowDefinition {  }
            };

            // WHEN
            var mat = Matrix.With()
                .ColsCount(10)
                .RowsCount(1)
                .Rows(rows)
                .Build();

            // THEN
            Check.That(mat.RowsDefinitions).IsEquivalentTo(rows);
            Check.That(mat.CountOfRows).IsEqualTo(1);
            Check.That(mat.CountOfColumns).IsEqualTo(10);
        }

        [Fact]
        public void Should_Build_Matrix_With_Values()
        {
            // GIVEN
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 123 }, { "colB", 234 } } }
            };

            // WHEN
            var mat = Matrix.With()
                .RowValues(values)
                .Build();

            // THEN
            Check.That(mat.RowValues).IsEquivalentTo(values);
            Check.That(mat.CountOfRows).IsEqualTo(1);
            Check.That(mat.CountOfColumns).IsEqualTo(2);
        }

        [Fact]
        public void Should_Build_Matrix_With_Cells_Without_Headers()
        {
            // GIVEN
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 11 }, { "colB", 44 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 22 }, { "colB", 55 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 33 }, { "colB", 66 } } },
            };

            // WHEN
            var mat = Matrix.With()
                .WithoutHeadersRow()
                .RowValues(values)
                .Build();

            // THEN
            Check.That(mat.RowValues).HasSize(3);
            List<MatrixCellValue> cells = mat.RowValues.SelectMany(rv => rv.Cells).ToList();
            Check.That(cells).HasSize(6);
            Check.That(cells.Extracting(nameof(MatrixCellValue.Value))).ContainsExactly(11, 44, 22, 55, 33, 66);
            Check.That(cells.Extracting(nameof(MatrixCellValue.RowIndex))).ContainsExactly(0, 0, 1, 1, 2, 2);
            Check.That(cells.Extracting(nameof(MatrixCellValue.ColIndex))).ContainsExactly(0, 1, 0, 1, 0, 1);
            Check.That(cells.Extracting(nameof(MatrixCellValue.Address))).ContainsExactly("A1", "B1", "A2", "B2", "A3", "B3");
        }

        [Fact]
        public void Should_Build_Matrix_With_Cells_With_Headers()
        {
            // GIVEN
            var cols = new List<ColumnDefinition> { 
                new ColumnDefinition { Name = "colA", Label = "I'm A" }, 
                new ColumnDefinition { Name = "colB", Label = "I'm B" } 
            };
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 11 }, { "colB", 44 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 22 }, { "colB", 55 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 33 }, { "colB", 66 } } },
            };

            // WHEN
            var mat = Matrix.With()
                .Cols(cols)
                .RowValues(values)
                .Build();

            // THEN
            Check.That(mat.RowValues).HasSize(3);
            List<MatrixCellValue> cells = mat.RowValues.SelectMany(rv => rv.Cells).ToList();
            Check.That(cells).HasSize(6);
            Check.That(cells.Extracting(nameof(MatrixCellValue.Value))).ContainsExactly(11, 44, 22, 55, 33, 66);
            Check.That(cells.Extracting(nameof(MatrixCellValue.RowIndex))).ContainsExactly(1, 1, 2, 2, 3, 3);
            Check.That(cells.Extracting(nameof(MatrixCellValue.ColIndex))).ContainsExactly(0, 1, 0, 1, 0, 1);
            Check.That(cells.Extracting(nameof(MatrixCellValue.Address))).ContainsExactly("A2", "B2", "A3", "B3", "A4", "B4");
        }

        [Fact]
        public void Should_Get_Row_From_Cell()
        {
            // GIVEN
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 11 }, { "colB", 44 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 22 }, { "colB", 55 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 33 }, { "colB", 66 } } },
            };

            var mat = Matrix.With()
                .RowValues(values)
                .Build();

            // WHEN
            RowValue row = mat.Row(values[0].Cells.ElementAt(0));

            // THEN
            Check.That(row).IsEqualTo(values[0]);
        }

        [Fact]
        public void Should_Get_Col_From_Row()
        {
            // GIVEN
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 11 }, { "colB", 44 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 22 }, { "colB", 55 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 33 }, { "colB", 66 } } },
            };

            var mat = Matrix.With()
                .RowValues(values)
                .Build();

            // WHEN
            MatrixCellValue cell = mat.Row(values[0].Cells.ElementAt(0)).Col("colB");

            // THEN
            Check.That(cell.RowIndex).IsEqualTo(1);
            Check.That(cell.ColIndex).IsEqualTo(1);
            Check.That(cell.Address).IsEqualTo("B2");
            Check.That(cell.Value).IsEqualTo(44);
        }

        [Fact]
        public void Should_Get_Col_From_Cell()
        {
            // GIVEN
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 11 }, { "colB", 44 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 22 }, { "colB", 55 } } },
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 33 }, { "colB", 66 } } },
            };

            var mat = Matrix.With()
                .RowValues(values)
                .Build();

            // WHEN
            ColumnCellReader colReader = mat.Col(values[0].Cells.ElementAt(0));

            // THEN
            Check.That(colReader.Values).ContainsExactly(11, 22, 33);
            Check.That(colReader.Cells.Extracting(nameof(MatrixCellValue.ColIndex))).IsOnlyMadeOf(0);
            Check.That(colReader.Cells.Extracting(nameof(MatrixCellValue.ColName))).IsOnlyMadeOf("colA");
            Check.That(colReader.Cells.Extracting(nameof(MatrixCellValue.Address))).ContainsExactly("A2", "A3", "A4");
        }

        public class TestDataColumnDefinition1 : IEnumerable<object[]>
        {
            public IEnumerator<object[]> GetEnumerator()
            {
                var dataList = GetData();
                foreach (var data in dataList)
                {
                    yield return new object[] { 
                        dataList,
                        data
                    };
                }
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            private static List<ColumnDefinition> GetData()
            {
                return new List<ColumnDefinition>
                {
                    new ColumnDefinition { Name = "ColA" },
                    new ColumnDefinition { Name = "ColB" },
                };
            }
        }

        [Theory]
        [ClassData(typeof(TestDataColumnDefinition1))]
        public void Should_Return_Column_Definition_By_Key(List<ColumnDefinition> cols, ColumnDefinition expectedColDef)
        {
            var mat = Matrix.With()
                .Cols(cols)
                .RowsCount(2)
                .Build();

            // WHEN
            ColumnDefinition colDef = mat.GetOwnColumnByIndex(expectedColDef.Index);

            // THEN
            Check.That(colDef).IsEqualTo(expectedColDef);
        }

        public class TestDataRowDefinition1 : IEnumerable<object[]>
        {
            public IEnumerator<object[]> GetEnumerator()
            {
                var dataList = GetData();
                foreach (var data in dataList)
                {
                    yield return new object[] {
                        dataList,
                        data
                    };
                }

                // Should return the Default RowDefinition when the specified key don't match any key
                yield return new object[]
                {
                    dataList,
                    new RowDefinition { Key = RowDefinition.DEFAULT_KEY }
                };
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            private static List<RowDefinition> GetData()
            {
                return new List<RowDefinition>
                {
                    new RowDefinition { Key = "RowA" },
                    new RowDefinition { Key = "RowB" },
                };
            }
        }

        [Theory]
        [ClassData(typeof(TestDataRowDefinition1))]
        public void Should_Return_Row_Definition_By_Key(List<RowDefinition> rows, RowDefinition expectedRowDef)
        {
            var mat = Matrix.With()
                .ColsCount(2)
                .RowsCount(1)
                .Rows(rows)
                .Build();

            // WHEN
            RowDefinition rowDef = mat.GetRowByKey(expectedRowDef.Key);

            // THEN
            Check.That(rowDef.Key).IsEqualTo(expectedRowDef.Key);
        }
    }
}
