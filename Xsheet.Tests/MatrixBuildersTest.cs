﻿using NFluent;
using NPOI.SS.Formula.Functions;
using System.Collections.Generic;
using System.Linq;
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
            var values = new List<RowValue>
            {
                new RowValue { ValuesByColName = new Dictionary<string, object>{ { "colA", 11 }, { "colB", 44 } } },
                new RowValue { Key = "ROW_1", ValuesByColName = new Dictionary<string, object>{ { "colA", 22 }, { "colB", 55 } } },
                new RowValue { Key = "ROW_2", ValuesByColName = new Dictionary<string, object>{ { "colA", 33 }, { "colB", 66 } } },
                new RowValue { Key = "ROW_3", ValuesByColName = new Dictionary<string, object>{ { "colA", 44 }, { "colB", 77 } } },
                new RowValue { Key = "ROW_4", ValuesByColName = new Dictionary<string, object>{ { "colA", 55 }, { "colB", 88 } } },
            };

            IFormat format = null;

            // WHEN
            var mat = Builders.New
                .Cols()
                    .Add(name: "colA", label: "I'm A", dataType: DataTypes.Boolean, headerCellFormat: format)
                    .Add(
                        new ColumnDefinition { Name = "colB", Label = "I'm B"}, 
                        new ColumnDefinition { Name = "colC", Label = "I'm C" }
                    )
                    .Add("colD", "I'm D")
                    .Continue()
                .Rows()
                    .Add(key: "ROW_1", defaultCellFormat: format)
                    .Add(new RowDefinition { Key = "ROW_2", DefaultCellFormat = format, FormatsByColName = null, ValuesMapping = null })
                        .AddFormat("ROW_3", format)
                        .AddValueMap("colB", (mat, cell) => "BBB")
                    .Add("ROW_4", format)
                    .Continue()
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
    }
}