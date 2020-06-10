using System.Collections.Generic;
using System.Linq;
using Xsheet;

namespace XSheet.Renderers.NPOI.Formats
{
    public class BasicFormat : IFormat
    {
        public string FontStyle { get; set; }
        public string BackgroundColor { get; set; }
        public int? FontSize { get; set; }

        public static BasicFormat MergeFormats(List<BasicFormat> formats)
        {
            return formats.Aggregate(new BasicFormat(), (mergedFormat, nextFormat) =>
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
