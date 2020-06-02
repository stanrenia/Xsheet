using System.Collections.Generic;
using System.Linq;

namespace Xsheet
{
    public class ColumnCellReader
    {
        private readonly IEnumerable<MatrixCellValue> _cells;

        public ColumnCellReader(IEnumerable<MatrixCellValue> cells)
        {
            _cells = cells;
        }

        public List<MatrixCellValue> Cells { get => _cells.ToList(); }
        public List<object> Values { get => Cells.Select(c => c.Value).ToList(); }

        public IEnumerable<MatrixCellValue> CellsOfRowKey(string rowKey)
        {
            return Cells.Where(c => c.Row.Key == "WEEK");
        }

        public List<MatrixCellValue> CellsBetween(MatrixCellValue startCell, int endRowIndex)
        {
            return CellsBetween(startCell.RowIndex, endRowIndex);
        }

        public List<MatrixCellValue> CellsBetween(int startRowIndex, MatrixCellValue endCell)
        {
            return CellsBetween(startRowIndex, endCell.RowIndex);
        }

        public List<MatrixCellValue> CellsBetween(int startRowIndex, int endRowIndex)
        {
            return Cells.Where(c => c.RowIndex < startRowIndex && c.RowIndex > endRowIndex)
                .ToList();
        }
    }

    public static class ColumnCellReaderExtensions
    {
        public static string Addresses(this IEnumerable<MatrixCellValue> cells, string separator = ",")
        {
            return string.Join(separator, cells.Select(c => c.Address));
        }
    }
}
