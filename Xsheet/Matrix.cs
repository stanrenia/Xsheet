using FluentValidation;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Xsheet
{
    public class Matrix
    {
        private IEnumerable<ColumnDefinition> _columnsDefinitions = new List<ColumnDefinition>();
        private Dictionary<string, ColumnDefinition> _columnsDefinitionsByKey = new Dictionary<string, ColumnDefinition>();

        public int CountOfColumns { get; internal set; }
        public int CountOfRows { get; internal set; }

        public IEnumerable<ColumnDefinition> ColumnsDefinitions
        {
            get => _columnsDefinitions;
            internal set
            {
                if (value != null)
                {

                    _columnsDefinitions = value;

                    _columnsDefinitionsByKey = _columnsDefinitions
                        .Select((col, i) =>
                        {
                            if (col.Index == ColumnDefinition.UNDEFINED_INDEX_VALUE)
                            {
                                col.Index = i;
                            }
                            return col;
                        })
                        .ToDictionary(col => col.Key, col => col);
                }
            }
        }

        public IEnumerable<RowDefinition> RowsDefinitions { get; internal set; } = new List<RowDefinition>();
        public IEnumerable<RowValue> RowValues { get; internal set; } = new List<RowValue>();
        public IEnumerable<MatrixCellValue> Values { get; internal set; } = new List<MatrixCellValue>();

        public bool HasHeaders { get => ColumnsDefinitions.Count() > 0; }

        private Matrix() { }

        public static MatrixBuilder With() => new MatrixBuilder();

        public Matrix ConcatX(Matrix mat)
        {
            throw new NotImplementedException();
        }

        public Matrix ConcatY(Matrix mat)
        {
            throw new NotImplementedException();
        }

        public ColumnDefinition GetColumnByKey(string key)
        {
            return _columnsDefinitionsByKey[key];
        }

        public class MatrixBuilder
        {
            private int CountOfRows;
            private int CountOfCols;
            private List<ColumnDefinition> ColumnsDefinition;
            private List<RowDefinition> RowsDefinition;
            private List<RowValue> _rowValues;

            public MatrixBuilder Dimensions(int rowsCount, int colsCount)
            {
                CountOfRows = rowsCount;
                CountOfCols = colsCount;
                return this;
            }

            public MatrixBuilder RowsCount(int rowsCount)
            {
                CountOfRows = rowsCount;
                return this;
            }

            public MatrixBuilder ColsCount(int colsCount)
            {
                CountOfCols = colsCount;
                return this;
            }

            public MatrixBuilder Cols(List<ColumnDefinition> cols)
            {
                ColumnsDefinition = cols;
                if (CountOfCols < cols.Count)
                {
                    CountOfCols = cols.Count;
                }
                return this;
            }

            public MatrixBuilder Rows(List<RowDefinition> rows)
            {
                RowsDefinition = rows;
                if (CountOfRows < rows.Count)
                {
                    CountOfRows = rows.Count;
                }
                return this;
            }

            public MatrixBuilder RowValues(List<RowValue> values)
            {
                _rowValues = values;
                var maxColCount = values.Max(rowValue => rowValue.ValuesByCol.Keys.Count);
                if (CountOfCols < maxColCount)
                {
                    CountOfCols = maxColCount;
                }
                if (CountOfRows < values.Count)
                {
                    CountOfRows = values.Count;
                }
                return this;
            }

            public Matrix Build()
            {
                var validator = new Validator();
                var mat = new Matrix
                {
                    CountOfRows = CountOfRows,
                    CountOfColumns = CountOfCols,
                    ColumnsDefinitions = ColumnsDefinition,
                    RowsDefinitions = RowsDefinition,
                    RowValues = _rowValues
                };

                var results = validator.Validate(mat);
                if (!results.IsValid)
                {
                    var err = results.Errors.First();
                    throw new ArgumentException($"One or more values are invalid: {err.ErrorMessage}");
                }

                return mat;
            }

            protected class Validator : AbstractValidator<Matrix>
            {
                public Validator()
                {
                    RuleFor(m => m.CountOfRows).GreaterThan(0);
                    RuleFor(m => m.CountOfColumns).GreaterThan(0);
                }
            }
        }
    }
}
