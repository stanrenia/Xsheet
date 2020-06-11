using System;
using System.Collections.Generic;

namespace Xsheet
{
    public class RowDefinition
    {
        public const string DEFAULT_KEY = "DEFAULT_KEY";
        public string Key { get; set; } = DEFAULT_KEY;
        public bool IsDefault { get => Key == DEFAULT_KEY; }
        public IFormat DefaultCellFormat { get; set; }
        public Dictionary<string, IFormat> FormatsByColName { get; set; } = new Dictionary<string, IFormat>();
        public Dictionary<int, IFormat> FormatsByColIndex { get; internal set; } = new Dictionary<int, IFormat>();
        public Dictionary<string, Func<Matrix, MatrixCellValue, object>> ValuesMapping { get; set; } = new Dictionary<string, Func<Matrix, MatrixCellValue, object>>();
    }
}
