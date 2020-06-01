﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace Xsheet
{
    public class RowValue
    {
        public string Key { get; set; }

        public Dictionary<string, object> ValuesByColName { get; set; } = new Dictionary<string, object>();

        public Dictionary<int, object> ValuesByColIndex { get; internal set; } = new Dictionary<int, object>();
        public IEnumerable<MatrixCellValue> Cells { get; internal set; } = new List<MatrixCellValue>();

        public MatrixCellValue Col(string colName)
        {
            return Cells.First(c => c.ColName == colName);
        }
    }
}
