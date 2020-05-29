using System.Collections.Generic;

namespace Xsheet
{
    public class RowValue
    {
        public string Key { get; set; }

        public Dictionary<string, object> ValuesByColName { get; set; } = new Dictionary<string, object>();

        public Dictionary<int, object> ValuesByColIndex { get; internal set; } = new Dictionary<int, object>();
    }
}
