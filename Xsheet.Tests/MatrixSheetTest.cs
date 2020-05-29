using NFluent;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
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
                        .Select(col => new { Col = $"ACol{col}", Value = $"AValue{col}{line}" })
                        .ToDictionary(o => o.Col, o => (object)o.Value)
                })
                .ToList();

            var values2 = Enumerable.Range(1, 3)
                .Select(line => new RowValue
                {
                    ValuesByColName = Enumerable.Range(1, 3)
                        .Select(col => new { Col = $"BCol{col}", Value = $"BValue{col}{line}" })
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
                return new { Key = col.Name, DataType = col.DataType, Index = i };
            });

            Check.That(m3.CountOfRows).Equals(3);
            Check.That(m3.CountOfColumns).Equals(7);
            Check.That(m3.ColumnsDefinitions).Considering().Properties.HasFieldsWithSameValues(expected);
            //Check.That(m3.GetColumnByKey())
        }
    }
}