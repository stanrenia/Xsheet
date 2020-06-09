using NFluent;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using XSheet.Renderers.Formats;
using Xunit;

namespace Xsheet.Tests
{
    public class MatrixSheetTest
    {
        public MatrixSheetTest()
        {

        }

        [Fact]
        public void Should_Add_Width_And_Keep_Height_When_Concat_Horizontally()
        {
            // GIVEN
            var m1 = Matrix.With().Key(index: 1)
                .Dimensions(2, 4)
                .Build();

            var m2 = Matrix.With().Key(index: 2)
                .Dimensions(3, 3)
                .Build();

            // WHEN
            var m3 = m1.ConcatX(m2);

            // THEN
            Check.That(m3.CountOfRows).Equals(3);
            Check.That(m3.CountOfColumns).Equals(7);
        }

        [Fact]
        public void Should_Add_Height_And_Keep_Width_When_Concat_Vertically()
        {
            // GIVEN
            var m1 = Matrix.With().Key(index: 1)
                .Dimensions(2, 4)
                .Build();

            var m2 = Matrix.With().Key(index: 2)
                .Dimensions(3, 3)
                .Build();

            // WHEN
            var m3 = m1.ConcatY(m2);

            // THEN
            Check.That(m3.CountOfRows).Equals(4);
            Check.That(m3.CountOfColumns).Equals(5);
        }

        [Fact]
        public void Should_Add_Width_And_Keep_Height_When_Concat_Horizontally_With_Values()
        {
            // GIVEN
            var values1 = Enumerable.Range(1, 3)
                .Select(line => new RowValue
                {
                    ValuesByColName = Enumerable.Range(1, 4)
                        .Select(col => new { Col = $"ACol{col}", Value = $"A{line}{col}" })
                        .ToDictionary(o => o.Col, o => (object)o.Value)
                })
                .ToList();

            var values2 = Enumerable.Range(1, 3)
                .Select(line => new RowValue
                {
                    ValuesByColName = Enumerable.Range(1, 3)
                        .Select(col => new { Col = $"BCol{col}", Value = $"B{line}{col}" })
                        .ToDictionary(o => o.Col, o => (object)o.Value)
                })
                .ToList();

            var m1 = Matrix.With().Key(index: 1)
                .RowValues(values1)
                .Build();

            var m2 = Matrix.With().Key(index: 2)
                .RowValues(values2)
                .Build();

            // WHEN
            var m3 = m1.ConcatX(m2);

            // THEN
            Check.That(m3.CountOfRows).Equals(3);
            Check.That(m3.CountOfColumns).Equals(7);
            var valuesM3 = m3.RowValues.ToList();
            Check.That(valuesM3[0].ValuesByColIndex.Keys.OrderBy(i => i)).ContainsExactly(0, 1, 2, 3, 4, 5, 6);
            Check.That(valuesM3[0].ValuesByColIndex.Values).Contains("A11", "A12", "A13", "A14", "B11", "B12", "B13");
            Check.That(valuesM3[1].ValuesByColIndex.Keys.OrderBy(i => i)).ContainsExactly(0, 1, 2, 3, 4, 5, 6);
            Check.That(valuesM3[1].ValuesByColIndex.Values).Contains("A21", "A22", "A23", "A24", "B21", "B22", "B23");
            Check.That(valuesM3[2].ValuesByColIndex.Keys.OrderBy(i => i)).ContainsExactly(0, 1, 2, 3, 4, 5, 6);
            Check.That(valuesM3[2].ValuesByColIndex.Values).Contains("A31", "A32", "A33", "A34", "B31", "B32", "B33");
        }

        [Fact]
        public void Should_Not_Concat_When_Matrices_Have_The_Same_Key()
        {
            var m1 = Matrix.With().Dimensions(2, 2).Build();
            var m2 = Matrix.With().Dimensions(2, 2).Build();

            Check.ThatCode(() => m1.ConcatX(m2)).Throws<InvalidOperationException>();
            Check.ThatCode(() => m1.ConcatY(m2)).Throws<InvalidOperationException>();
        }

        [Fact]
        public void Should_Add_ColumnDefinitions_When_Concat_Horizontally()
        {
            // GIVEN
            var colDefs = Enumerable.Range(1, 3)
                .Select(col => new ColumnDefinition
                {
                    Name = $"Col{col}",
                    DataType = DataTypes.Text
                })
                .ToList();

            var colDefs2 = Enumerable.Range(1, 4)
                .Select(col => new ColumnDefinition
                {
                    Name = $"Col{col}",
                    DataType = DataTypes.Number
                })
                .ToList();

            var m1 = Matrix.With().Key(index: 1)
                .Cols(colDefs)
                .RowsCount(2)
                .Build();

            var m2 = Matrix.With().Key(index: 2)
                .Cols(colDefs2)
                .RowsCount(3)
                .Build();

            // WHEN
            var m3 = m1.ConcatX(m2);

            // THEN
            var expected = colDefs.Concat(colDefs2).Select((col, i) =>
            {
                return new { Name = col.Name, DataType = col.DataType, Index = i };
            });

            var keys1 = colDefs.Select(col => new ColumnKey(new MatrixKey("", 1), col.Index, col.Name));
            var keys2 = colDefs2.Select(col => new ColumnKey(new MatrixKey("", 2), col.Index, col.Name));
            var expectedKeys = keys1.Concat(keys2).ToList();

            Check.That(m3.CountOfRows).Equals(3);
            Check.That(m3.CountOfColumns).Equals(7);
            Check.That(m3.ColumnsDefinitions.Extracting(nameof(ColumnDefinition.Index))).ContainsExactly(expected.Select(e => e.Index));
            Check.That(m3.ColumnsDefinitions.Extracting(nameof(ColumnDefinition.Name))).ContainsExactly(expected.Select(e => e.Name));
            Check.That(m3.ColumnsDefinitions.Extracting(nameof(ColumnDefinition.DataType))).ContainsExactly(expected.Select(e => e.DataType));
            Check.That(m3.ColumnsDefinitions.Extracting(nameof(ColumnDefinition.Key))).ContainsExactly(expectedKeys);
        }

        [Fact]
        public void Should_Add_RowDefinitions_When_Concat_Horizontally()
        {
            // GIVEN
            // Col1 -> Col3
            var colDefs = Enumerable.Range(1, 2)
                .Select(col => new ColumnDefinition
                {
                    Name = $"Col{col}",
                    DataType = DataTypes.Text
                })
                .ToList();

            var colDefs2 = Enumerable.Range(1, 3)
                .Select(col => new ColumnDefinition
                {
                    Name = $"Col{col}",
                    DataType = DataTypes.Text
                })
                .ToList();

            // R1 -> R2
            var rowDefs = Enumerable.Range(1, 2)
                .Select(i => new RowDefinition
                {
                    Key = $"R{i}",
                    FormatsByColName = new Dictionary<string, IFormat> {
                        { "Col1", new BasicFormat { FontSize = 10 } },
                        { "Col2", new BasicFormat { FontSize = 11 } }
                    }
                })
                .ToList();

            // R2 -> R3
            var rowDefs2 = Enumerable.Range(2, 2)
                .Select(i => new RowDefinition
                {
                    Key = $"R{i}",
                    FormatsByColName = new Dictionary<string, IFormat> {
                        { "Col1", new BasicFormat { FontSize = 20 } },
                        { "Col2", new BasicFormat { FontSize = 21 } },
                        { "Col3", new BasicFormat { FontSize = 22 } }
                    }
                })
                .ToList();

            var m1 = Matrix.With().Key(index: 1)
                .Cols(colDefs)
                .Rows(rowDefs)
                .RowsCount(5)
                .Build();

            var m2 = Matrix.With().Key(index: 2)
                .Cols(colDefs2)
                .Rows(rowDefs2)
                .RowsCount(9)
                .Build();

            // WHEN
            var m3 = m1.ConcatX(m2);

            // THEN
            Check.That(m3.CountOfRows).Equals(9);
            Check.That(m3.CountOfColumns).Equals(5);
            Check.That(m3.RowsDefinitions.Extracting(nameof(RowDefinition.Key))).ContainsExactly("R1", "R2", "R3");
            // Check R1
            var formatsR1 = m3.RowsDefinitions.First(r => r.Key == "R1").FormatsByColIndex;
            Check.That(formatsR1.Keys).ContainsExactly(0, 1);
            Check.That(formatsR1.Values.Cast<BasicFormat>().Extracting(nameof(BasicFormat.FontSize))).ContainsExactly(10, 11);
            // Check R2
            var formatsR2 = m3.RowsDefinitions.First(r => r.Key == "R2").FormatsByColIndex;
            Check.That(formatsR2.Keys).ContainsExactly(0, 1, 2, 3, 4);
            Check.That(formatsR2.Values.Cast<BasicFormat>().Extracting(nameof(BasicFormat.FontSize))).ContainsExactly(10, 11, 20, 21, 22);
            // Check R3
            var formatsR3 = m3.RowsDefinitions.First(r => r.Key == "R3").FormatsByColIndex;
            Check.That(formatsR3.Keys).ContainsExactly(2, 3, 4);
            Check.That(formatsR3.Values.Cast<BasicFormat>().Extracting(nameof(BasicFormat.FontSize))).ContainsExactly(20, 21, 22);
        }

        [Fact]
        public void Should_Throw_Exception_When_Conact_RowDefinitions_With_Same_Key_And_Strategy_Is_RaiseError()
        {
            // GIVEN
            var colDefs = Enumerable.Range(1, 3).Select(col => new ColumnDefinition{Name = $"Col{col}"}).ToList();
            var colDefs2 = Enumerable.Range(1, 3).Select(col => new ColumnDefinition{Name = $"Col{col}"}).ToList();

            var rowDefs = Enumerable.Range(1, 2).Select(i => new RowDefinition { Key = $"R{i}" }).ToList();
            var rowDefs2 = Enumerable.Range(1, 2).Select(i => new RowDefinition { Key = $"R{i}" }).ToList();

            var m1 = Matrix.With().Key(index: 1)
                .Cols(colDefs)
                .Rows(rowDefs)
                .RowsCount(5)
                .Build();

            var m2 = Matrix.With().Key(index: 2)
                .Cols(colDefs2)
                .Rows(rowDefs2)
                .RowsCount(9)
                .Build();

            // WHEN
            Action action = () => m1.ConcatX(m2, MatrixConcatStrategy.RaiseError);

            // THEN
            Check.ThatCode(action).Throws<InvalidOperationException>();
        }
    }
}