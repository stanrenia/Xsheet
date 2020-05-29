using System.Collections.Generic;

namespace Xsheet
{
    public class RowDefinition
    {
        public const string DEFAULT_KEY = "DEFAULT_KEY";
        public string Key { get; set; } = DEFAULT_KEY;
        public bool IsDefault { get => Key == DEFAULT_KEY; }
        public Format DefaultCellFormat { get; set; }
        public Dictionary<string, Format> FormatsByColName { get; set; } = new Dictionary<string, Format>();
        public Dictionary<int, Format> FormatsByColIndex { get; internal set; } = new Dictionary<int, Format>();
    }
}