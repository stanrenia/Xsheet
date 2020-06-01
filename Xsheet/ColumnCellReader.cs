using System.Collections.Generic;
using System.Linq;

namespace Xsheet
{
    public class ColumnCellReader
    {
        public ColumnCellReader(IEnumerable<MatrixCellValue> cells)
        {
            Cells = cells.ToList();
        }

        public List<MatrixCellValue> Cells { get; }
        public List<object> Values { get => Cells.Select(c => c.Value).ToList(); }
    }
}
