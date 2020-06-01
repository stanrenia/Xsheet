namespace Xsheet
{
    public class MatrixCellValue
    {
        public MatrixCellValue(int rowIndex, int colIndex, string colName, object value)
        {
            RowIndex = rowIndex;
            ColIndex = colIndex;
            ColName = colName;
            Value = value;
            Address = $"{CellAddressCalculator.GetColumnLetters(colIndex)}{rowIndex + 1}";
        }

        public int RowIndex { get; }
        public int ColIndex { get; }
        public string ColName { get; }
        public string Address { get; }
        public object Value { get; }
    }
}