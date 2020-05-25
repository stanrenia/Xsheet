using NFluent;
using System.Collections;
using System.Collections.Generic;
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
                new ColumnDefinition { Key = "ColA" }
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
                new RowValue { ValuesByCol = new Dictionary<string, object>{ { "colA", 123 }, { "colB", 234 } } }
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
                    new ColumnDefinition { Key = "ColA" },
                    new ColumnDefinition { Key = "ColB" },
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
            ColumnDefinition colDef = mat.GetColumnByKey(expectedColDef.Key);

            // THEN
            Check.That(colDef).IsEqualTo(expectedColDef);
        }
    }
}
