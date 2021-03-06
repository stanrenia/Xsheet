﻿namespace Xsheet
{
    public class MatrixCellValue
    {
        public MatrixCellValue(MatrixKey key, RowValue row, int rowIndex, int colIndex, string colName, object value)
        {
            MatrixKey = key;
            Row = row;
            RowIndex = rowIndex;
            ColIndex = colIndex;
            ColName = colName;
            Value = value;
            Address = $"{CellAddressCalculator.GetColumnLetters(colIndex)}{rowIndex + 1}";
        }

        public MatrixKey MatrixKey { get; }
        public int RowIndex { get; }
        public int ColIndex { get; }
        public string ColName { get; }
        public string Address { get; }
        public object Value { get; }

        public RowValue Row { get; }
    }
}
