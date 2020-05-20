using System;
using System.Collections.Generic;

namespace Xsheet
{
    public class Matrix
    {
        public short CountOfColumns { get; set; }
        public short CountOfRows { get; set; }

        public IEnumerable<ColumnDefinition> ColumnsDefinitions { get; set; } = new List<ColumnDefinition>();
        public IEnumerable<RowDefinition> RowsDefinitions { get; set; } = new List<RowDefinition>();
        public IEnumerable<MatrixCellValue> Values { get; set; } = new List<MatrixCellValue>();

        public Matrix ConcatX(Matrix mat)
        {
            throw new NotImplementedException();
        }

        public Matrix ConcatY(Matrix mat)
        {
            throw new NotImplementedException();
        }
    }
}
