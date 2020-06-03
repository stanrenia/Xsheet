using System.Linq;

namespace Xsheet
{
    public class RowCellReader
    {
        private readonly Matrix _matrix;
        private readonly MatrixCellValue _cell;
        private readonly int _cellRowIndex;

        public RowCellReader(Matrix matrix, MatrixCellValue cell, int cellRowIndex)
        {
            _matrix = matrix;
            _cell = cell;
            _cellRowIndex = cellRowIndex;
        }

        public MatrixCellValue Col(string colName)
        {
            var colDef = _matrix.ColumnsDefinitions
                .Where(c => c.Name == colName && c.Key.MatrixKey.Equals(_cell.MatrixKey))
                .First();
            return _matrix.RowValues
                .ElementAt(_matrix.HasHeaders ? _cellRowIndex - 1 : _cellRowIndex)
                .Cells.First(c => c.ColIndex == colDef.Index);
        }


    }
}
