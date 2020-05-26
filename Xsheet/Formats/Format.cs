using System.Collections.Generic;
using System.Linq;

namespace Xsheet
{
    public class Format
    {
        public string FontStyle { get; set; }
        public string BackgroundColor { get; set; }
        public int? FontSize { get; set; }

        public static Format MergeFormats(List<Format> formats)
        {
            return formats.Aggregate(new Format(), (mergedFormat, nextFormat) =>
            {
                if (nextFormat.FontSize.HasValue)
                {
                    mergedFormat.FontSize = nextFormat.FontSize;
                }
                if (!string.IsNullOrEmpty(nextFormat.FontStyle))
                {
                    mergedFormat.FontStyle = nextFormat.FontStyle;
                }
                if (!string.IsNullOrEmpty(nextFormat.BackgroundColor))
                {
                    mergedFormat.BackgroundColor = nextFormat.BackgroundColor;
                }
                return mergedFormat;
            });
        }
    }
}
