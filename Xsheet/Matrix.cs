using FluentValidation;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Xsheet
{
    public class Matrix
    {
        private IEnumerable<ColumnDefinition> _columnsDefinitions = new List<ColumnDefinition>();
        private Dictionary<int, ColumnDefinition> _columnsDefinitionsByIndex = new Dictionary<int, ColumnDefinition>();
        private Dictionary<ColumnKey, ColumnDefinition> _columnsDefinitionsByKey = new Dictionary<ColumnKey, ColumnDefinition>();
        private IEnumerable<RowDefinition> _rowsDefinitions = new List<RowDefinition>();
        private Dictionary<string, RowDefinition> _rowsDefinitionsByKey = new Dictionary<string, RowDefinition>();
        private IEnumerable<RowValue> _rowValues = new List<RowValue>();

        public MatrixKey Key { get; internal set; }

        public int CountOfColumns { get; internal set; }
        public int CountOfRows { get; internal set; }

        public IEnumerable<ColumnDefinition> ColumnsDefinitions
        {
            get => _columnsDefinitions;
            internal set
            {
                if (value != null)
                {
                    if (value.Any(colDef => colDef.Name == null))
                    {
                        throw new ArgumentNullException($"{nameof(ColumnDefinition)}.{nameof(ColumnDefinition.Name)} cannot be null");
                    }

                    _columnsDefinitions = value;

                    var colsWithIndex = _columnsDefinitions
                        .Select((col, i) =>
                        {
                            if (col.Index == ColumnDefinition.UNDEFINED_INDEX_VALUE)
                            {
                                col.Index = i;
                            }
                            return col;
                        });

                    _columnsDefinitionsByIndex = colsWithIndex.ToDictionary(col => col.Index);

                    _columnsDefinitionsByKey = colsWithIndex.ToDictionary(col => new ColumnKey(Key, col.Index, col.Name));
                }
            }
        }

        public IEnumerable<RowDefinition> RowsDefinitions
        {
            get => _rowsDefinitions;
            internal set
            {
                if (value != null)
                {
                    _rowsDefinitions = value;
                    _rowsDefinitionsByKey = _rowsDefinitions
                        .ToDictionary(row => row.Key, row => row);

                    // Build the formats by column index Dictionnary if not built
                    // Then it is altered only during concatenation
                    foreach (var rowDef in _rowsDefinitions
                        .Where(rowDef => rowDef.FormatsByColIndex.Count == 0 && rowDef.FormatsByColName.Count > 0))
                    {
                        rowDef.FormatsByColIndex = rowDef.FormatsByColName.ToDictionary(kv => GetOwnColumnIndex(kv.Key), kv => kv.Value);
                    }
                }
            }
        }

        private int GetOwnColumnIndex(string colName)
        {
            return _columnsDefinitions.Where(colDef => colDef.Name == colName).First().Index;
        }

        public IEnumerable<RowValue> RowValues
        {
            get => _rowValues;
            internal set
            {
                if (value != null)
                {
                    if (_columnsDefinitions.Count() == 0 && value.Count() > 0)
                    {
                        ColumnsDefinitions = value.SelectMany(rv => rv.ValuesByColName.Keys)
                            .Distinct()
                            .Select(key => new ColumnDefinition { Name = key })
                            .ToList();
                    }

                    _rowValues = value;

                    int rowIndex = HasHeaders ? 1 : 0;
                    // Build the values by column index Dictionnary if not built
                    // Then it is altered only during concatenation
                    foreach (var rowValue in _rowValues
                        .Where(rowValue => rowValue.ValuesByColIndex.Count == 0))
                    {
                        var cells = new List<MatrixCellValue>(CountOfColumns);
                        foreach (int colIndex in Enumerable.Range(0, CountOfColumns))
                        {
                            var cellValue = rowValue.ValuesByColName
                                .Where(aKV => GetOwnColumnIndex(aKV.Key) == colIndex)
                                .Select(aKV => new { Value = aKV.Value })
                                .FirstOrDefault();

                            var colName = GetOwnColumnByIndex(colIndex).Name;
                            rowValue.ValuesByColIndex.Add(colIndex, cellValue?.Value);
                            cells.Add(new MatrixCellValue(rowIndex, colIndex, colName, cellValue?.Value));
                        }
                        rowValue.Cells = cells;
                        rowIndex++;
                    }
                }
            }
        }

        public IEnumerable<MatrixCellValue> Values { get; internal set; } = new List<MatrixCellValue>();

        public bool HasHeaders { get => WithHeadersRow && ColumnsDefinitions.Count() > 0; }
        public bool WithHeadersRow { get; private set; }

        private Matrix() { }

        public static MatrixBuilder With() => new MatrixBuilder();

        // TODO What to do if rows defs contains same Key ?
        // -- May allow to define the behaviour like: 
        // -- Keep left/right, merge, raiseError
        public Matrix ConcatX(Matrix aMat, MatrixConcatStrategy rowsStrategy = MatrixConcatStrategy.KeepLeft)
        {
            if (Key.Equals(aMat.Key))
            {
                throw new InvalidOperationException($"Cannot concat Matricies with the same Key {Key}");
            }

            var leftColsCount = ColumnsDefinitions.Count();

            List<ColumnDefinition> cols = ConcatXColumnsDefinitions(aMat, leftColsCount);
            List<RowDefinition> rows = ConcatXRowsDefinitions(aMat, rowsStrategy, leftColsCount);
            List<RowValue> values = ConcatXRowValues(aMat, leftColsCount);

            var builder = Matrix.With()
                .Dimensions(Math.Max(this.CountOfRows, aMat.CountOfRows), this.CountOfColumns + aMat.CountOfColumns);

            if (cols.Count() > 0)
            {
                builder.Cols(cols);
            }
            if (rows.Count() > 0)
            {
                builder.Rows(rows);
            }
            if (values.Count() > 0)
            {
                builder.RowValues(values);
            }

            return builder.Build();
        }

        private List<ColumnDefinition> ConcatXColumnsDefinitions(Matrix rightMat, int leftColsCount)
        {
            return ColumnsDefinitions.Concat(rightMat.ColumnsDefinitions.Select(col =>
            {
                col.Index += leftColsCount;
                return col;
            })).ToList();
        }

        private List<RowDefinition> ConcatXRowsDefinitions(Matrix rightMat, MatrixConcatStrategy rowsStrategy, int leftColsCount)
        {
            var rows = new List<RowDefinition>();

            if (rowsStrategy == MatrixConcatStrategy.RaiseError
                && RowsDefinitions.Any(r => rightMat.RowsDefinitions.Any(r2 => r2.Key == r.Key)))
            {
                throw new InvalidOperationException($@"Cannot ConcatX using strategy {nameof(MatrixConcatStrategy.RaiseError)}:
                Some RowDefinitions have the same Key.");
            }
            else if (rowsStrategy == MatrixConcatStrategy.KeepRight)
            {
                throw new NotImplementedException("KeepRight is not supported yet");
            }
            else if (rowsStrategy == MatrixConcatStrategy.Merge)
            {
                throw new NotImplementedException("Merge is not supported yet");
            }

            var rightMatRowDefs = rightMat.RowsDefinitions.Select(row =>
            {
                row.FormatsByColIndex = row.FormatsByColIndex
                    .ToDictionary(kv => kv.Key + leftColsCount, kv => kv.Value);
                return row;
            });

            if (rowsStrategy == MatrixConcatStrategy.KeepLeft)
            {
                var rowDefstoAdd = new List<RowDefinition>();
                foreach (var rightRowDef in rightMatRowDefs)
                {
                    var leftRowDef = RowsDefinitions.FirstOrDefault(r => r.Key == rightRowDef.Key);
                    if (leftRowDef is null)
                    {
                        rowDefstoAdd.Add(rightRowDef);
                    }
                    else
                    {
                        if (this.CountOfColumns < rightMat.CountOfColumns)
                        {
                            foreach (var kv in rightRowDef.FormatsByColIndex.Where(kv => kv.Key > this.CountOfColumns - 1))
                            {
                                leftRowDef.FormatsByColIndex.Add(kv.Key, kv.Value);
                            }
                        }
                    }
                }

                rows = RowsDefinitions.Concat(rowDefstoAdd).ToList();
            }

            return rows;
        }

        public ColumnCellReader Col(MatrixCellValue cell)
        {
            return new ColumnCellReader(RowValues.SelectMany(rv => rv.Cells.Where(c => c.ColIndex == cell.ColIndex)));
        }

        public RowValue Row(MatrixCellValue cell)
        {
            return Row(cell.RowIndex);
        }

        public RowValue Row(int rowIndex)
        {
            return RowValues.ElementAt(rowIndex - (HasHeaders ? 1 : 0));
        }

        private List<RowValue> ConcatXRowValues(Matrix rightMat, int leftColsCount)
        {
            var values = this.RowValues.Select((leftValue, rowIndex) =>
            {
                var rightValue = rightMat.RowValues.ElementAtOrDefault(rowIndex);
                if (rightValue != null)
                {
                    leftValue.ValuesByColIndex = rightValue.ValuesByColIndex
                        .ToDictionary(kv => kv.Key + leftColsCount, kv => kv.Value)
                        .Concat(leftValue.ValuesByColIndex)
                        .ToDictionary(kv => kv.Key, kv => kv.Value);

                    leftValue.Cells = leftValue.Cells
                        .Concat(rightValue.Cells.Select(cell => new MatrixCellValue(
                            rowIndex, 
                            cell.ColIndex + leftColsCount, 
                            rightMat.GetOwnColumnByIndex(cell.ColIndex).Name, 
                            cell.Value
                        )))
                        .ToList();
                }

                return leftValue;
            });

            if (this.CountOfRows < rightMat.CountOfRows)
            {
                values = values.Concat(rightMat.RowValues.Skip(this.CountOfRows));
            }

            return values.ToList();
        }

        public Matrix ConcatY(Matrix aMat)
        {
            if (Key.Equals(aMat.Key))
            {
                throw new InvalidOperationException($"Cannot concat Matricies with the same Key {Key}");
            }

            return Matrix.With()
                .Dimensions(Math.Max(this.CountOfColumns, aMat.CountOfColumns), this.CountOfRows + aMat.CountOfRows)
                .Build();
        }

        public ColumnDefinition GetOwnColumnByIndex(int index)
        {
            return _columnsDefinitionsByIndex[index];
        }

        public RowDefinition GetRowByKey(string aKey)
        {
            var key = aKey is null ? RowDefinition.DEFAULT_KEY : aKey;
            _rowsDefinitionsByKey.TryGetValue(key, out RowDefinition row);

            return row is null ? GetDefaultRowDefinition() : row;
        }

        private RowDefinition GetDefaultRowDefinition()
        {
            return new RowDefinition { Key = RowDefinition.DEFAULT_KEY };
        }

        public class MatrixBuilder
        {
            private MatrixKey _key;
            private int CountOfRows;
            private int CountOfCols;
            private List<ColumnDefinition> ColumnsDefinition;
            private List<RowDefinition> RowsDefinition;
            private List<RowValue> _rowValues;
            private bool _withHeadersRow = true;

            public MatrixBuilder Key(MatrixKey key)
            {
                _key = key;
                return this;
            }
            public MatrixBuilder Key(string key = "", int index = 0)
            {
                _key = new MatrixKey(key, index);
                return this;
            }


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
                return this;
            }

            public MatrixBuilder RowValues(List<RowValue> values)
            {
                _rowValues = values;
                var maxColCount = values.Max(rowValue => rowValue.ValuesByColName.Keys.Count);
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
                    Key = _key,
                    WithHeadersRow = _withHeadersRow,
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

            public MatrixBuilder WithoutHeadersRow()
            {
                _withHeadersRow = false;
                return this;
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
