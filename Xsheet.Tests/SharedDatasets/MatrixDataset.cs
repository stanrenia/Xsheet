using System.Collections.Generic;

namespace Xsheet.Tests.SharedDatasets
{
    public class MatrixDataset
    {
        public List<ColumnDefinition> Cols { get; set; }
        public List<RowDefinition> Rows { get; set; }
        public List<RowValue> Values { get; set; }
        public Matrix Matrix { get; set; }
    }
}
