using System.Collections.Generic;

namespace Xsheet
{
    public class RowValue
    {
        public Dictionary<string, object> ValuesByCol { get; set; } = new Dictionary<string, object>();
    }
}
