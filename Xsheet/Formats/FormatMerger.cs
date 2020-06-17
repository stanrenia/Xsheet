using System.Collections.Generic;

namespace Xsheet.Formats
{
    public class FormatMerger
    {
        /// 1) Add default format of default row 
        /// 2) Add each Col format of default row
        /// 3) Add default format of current row
        /// 4) Add each Col format of current row
        public static List<T> GetFormatsOrderedByLessPrioritary<T>(RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex)
        {
            var formats = new List<T>();
            void AddFormatIfNotNull(IFormat f)
            {
                if (f != null && f is T concreteFormat)
                {
                    formats.Add(concreteFormat);
                }
            }

            AddFormatIfNotNull(defaultRowDef.DefaultCellFormat);
            defaultRowDef.FormatsByColIndex.TryGetValue(columnIndex, out IFormat defaultColFormat);
            AddFormatIfNotNull(defaultColFormat);
            AddFormatIfNotNull(rowDef.DefaultCellFormat);
            rowDef.FormatsByColIndex.TryGetValue(columnIndex, out IFormat colFormat);
            AddFormatIfNotNull(colFormat);
            return formats;
        }
    }
}
