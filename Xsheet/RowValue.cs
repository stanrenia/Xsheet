using System.Collections.Generic;

namespace Xsheet
{
    public class RowValue
    {
        public string Key { get; set; }

        public Dictionary<string, object> ValuesByCol { get; set; } = new Dictionary<string, object>();
    }
}
