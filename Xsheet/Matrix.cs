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

                    _columnsDefinitions = value.ToList();

                    var colsWithIndex = _columnsDefinitions
                        .Select((col, i) =>
                        {
                            if (col.Index == ColumnDefinition.UNDEFINED_INDEX_VALUE)
                            {
                                col.Index = i;
                            }

                            if (col.Key.Equals(default(ColumnKey)))
                            {
                                col.Key = new ColumnKey(Key, col.Index, col.Name);
                            }

                            return col;
                        });

                    _columnsDefinitionsByIndex = colsWithIndex.ToDictionary(col => col.Index);
                    _columnsDefinitionsByKey = colsWithIndex.ToDictionary(col => col.Key);
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
                    _rowsDefinitions = value.ToList();
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
                            cells.Add(new MatrixCellValue(this.Key, rowValue, rowIndex, colIndex, colName, cellValue?.Value));
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

        public static IMatrixBuilder With() => new Builder();

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
            var cols = ColumnsDefinitions.Concat(rightMat.ColumnsDefinitions.Select(col =>
            {
                col.Index += leftColsCount;
                col.Key = new ColumnKey(col.Key.MatrixKey, col.Index, col.Key.Key);
                return col;
            })).ToList();

            return cols;
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

        public RowCellReader Row(MatrixCellValue cell, int rowIndex = -1)
        {
            if (rowIndex == -1)
            {
                rowIndex = cell.RowIndex;
            }
            return new RowCellReader(this, cell, rowIndex);
        }

        public RowCellReader RowAbove(MatrixCellValue cell)
        {
            return new RowCellReader(this, cell, cell.RowIndex - 1);
        }

        public int RowIndexOfPrevious(string rowKey, MatrixCellValue startCell)
        {
            return RowIndexOfPrevious(rowKey, startCell.RowIndex);
        }

        public int RowIndexOfPrevious(string rowKey, int startRowIndex)
        {
            return this.RowValues
                .OrderByDescending(rv => rv.RowIndex)
                .FirstOrDefault(rv => rv.Key == rowKey && rv.RowIndex < startRowIndex)?.RowIndex ?? -1;
        }


        private List<RowValue> ConcatXRowValues(Matrix rightMat, int leftColsCount)
        {
            var values = this.RowValues.Select((leftValue, rowListIndex) =>
            {
                var rightValue = rightMat.RowValues.ElementAtOrDefault(rowListIndex);
                if (rightValue != null)
                {
                    leftValue.ValuesByColIndex = rightValue.ValuesByColIndex
                        .ToDictionary(kv => kv.Key + leftColsCount, kv => kv.Value)
                        .Concat(leftValue.ValuesByColIndex)
                        .ToDictionary(kv => kv.Key, kv => kv.Value);

                    leftValue.Cells = leftValue.Cells
                        .Concat(rightValue.Cells.Select(cell => new MatrixCellValue(
                            rightMat.Key,
                            rightValue,
                            leftValue.Cells.ElementAt(0).RowIndex, 
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

        public class Builder : IMatrixBuilder, IColsBuilder, IRowsBuilder, IRowBuilder
        {
            private List<ColumnDefinition> _cols = new List<ColumnDefinition>();
            private List<RowDefinition> _rows = new List<RowDefinition>();
            private List<RowValue> _values = new List<RowValue>();
            private int _countOfCols;
            private int _countOfRows;
            private bool _withHeadersRow = true;
            private MatrixKey _key;

            public IMatrixBuilder Key(string key = "", int index = 0)
            {
                _key = new MatrixKey(key, index);
                return this;
            }

            public IMatrixBuilder Dimensions(int rowsCount, int colsCount)
            {
                _countOfRows = rowsCount;
                _countOfCols = colsCount;
                return this;
            }

            public IMatrixBuilder RowsCount(int rowsCount)
            {
                _countOfRows = rowsCount;
                return this;
            }

            public IMatrixBuilder ColsCount(int colsCount)
            {
                _countOfCols = colsCount;
                return this;
            }

            public IColsBuilder Cols()
            {
                return this;
            }

            public IMatrixBuilder Cols(List<ColumnDefinition> cols)
            {
                _cols = cols;
                return this;
            }

            public IColsBuilder Col(string name = null, string label = null, DataTypes dataType = DataTypes.Text, IFormat headerCellFormat = null)
            {
                _cols.Add(new ColumnDefinition
                {
                    Name = name,
                    Label = label,
                    DataType = dataType,
                    HeaderCellFormat = headerCellFormat
                });
                return this;
            }

            public IRowsBuilder Rows()
            {
                return this;
            }

            public IMatrixBuilder Rows(List<RowDefinition> rows)
            {
                _rows = rows;
                return this;
            }

            IRowsBuilder IColsBuilder.Rows(List<RowDefinition> rows)
            {
                _rows = rows;
                return this;
            }

            private RowDefinition MakeRowDefinition(string key, IFormat defaultCellFormat)
            {
                return new RowDefinition
                {
                    Key = key is null ? RowDefinition.DEFAULT_KEY : key,
                    DefaultCellFormat = defaultCellFormat
                };
            }

            public IRowBuilder Row(string key = null, IFormat defaultCellFormat = null)
            {
                _rows.Add(MakeRowDefinition(key, defaultCellFormat));
                return this;
            }

            public IRowBuilder Format(string colName, IFormat format)
            {
                var rowDef = _rows.Last();
                rowDef.FormatsByColName.Add(colName, format);
                return this;
            }

            public IRowBuilder ValueMap(string colName, Func<Matrix, MatrixCellValue, object> func)
            {
                var rowDef = _rows.Last();
                rowDef.ValuesMapping.Add(colName, func);
                return this;
            }

            public IMatrixBuilder RowValues(IEnumerable<RowValue> rowValues)
            {
                _values = rowValues.ToList();
                return this;
            }

            public IMatrixBuilder WithoutHeadersRow()
            {
                _withHeadersRow = false;
                return this;
            }

            public Matrix Build()
            {
                if (_values.Count > 0)
                {
                    var maxColCount = _values.Max(rowValue => rowValue.ValuesByColName.Keys.Count);
                    if (_countOfCols < maxColCount)
                    {
                        _countOfCols = maxColCount;
                    }
                }

                if (_countOfRows < _values.Count)
                {
                    _countOfRows = _values.Count;
                }

                if (_countOfCols < _cols.Count)
                {
                    _countOfCols = _cols.Count;
                }

                var mat = new Matrix {
                    Key = _key,
                    CountOfColumns = _countOfCols,
                    CountOfRows = _countOfRows,
                    WithHeadersRow = _withHeadersRow,
                    ColumnsDefinitions = _cols,
                    RowsDefinitions = _rows,
                    RowValues = _values
                };

                var results = new Validator().Validate(mat);
                if (!results.IsValid)
                {
                    var err = results.Errors.First();
                    throw new ArgumentException($"One or more values are invalid: {err.ErrorMessage}");
                }

                return mat;
            }
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
