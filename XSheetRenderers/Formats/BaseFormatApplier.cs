using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using Xsheet;

namespace XSheet.Renderers.Formats
{
    public abstract class BaseFormatApplier<T> : FormatApplier
        where T : class, Format
    {
        /// <summary>
        /// Add the following Formats in order, then call MergeFormats()
        /// 1) Add default format of default row 
        /// 2) Add each Col format of default row
        /// 3) Add default format of current row
        /// 4) Add each Col format of current row
        /// Override MergeFormats() to change the default behavior, which is only taking the 4th format (the last one)
        /// </summary>
        public void ApplyFormatToCell(IWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex, ICell cell)
        {
            var formats = new List<T>();
            void AddFormatIfNotNull(Format f)
            {
                if (f != null && f is T nPOIFormat)
                {
                    formats.Add(nPOIFormat);
                }
            }

            AddFormatIfNotNull(defaultRowDef.DefaultCellFormat);
            defaultRowDef.FormatsByColIndex.TryGetValue(columnIndex, out Format defaultColFormat);
            AddFormatIfNotNull(defaultColFormat);
            AddFormatIfNotNull(rowDef.DefaultCellFormat);
            rowDef.FormatsByColIndex.TryGetValue(columnIndex, out Format colFormat);
            AddFormatIfNotNull(colFormat);

            var mergedFormat = MergeFormats(formats);

            if (mergedFormat != null)
            {
                ApplyFormatToCell(wb, cell, mergedFormat);
            }
        }

        public abstract void ApplyFormatToCell(IWorkbook wb, ICell cell, T format);

        public void ApplyFormatToCell(IWorkbook wb, ICell cell, Format format)
        {
            if (format != null)
            {
                if (format is T concreteFormat)
                {
                    ApplyFormatToCell(wb, cell, concreteFormat);
                }
                else
                {
                    throw new ArgumentException($"Given Format is not of Type {typeof(T)}, actual {format.GetType()}");
                }
            }
        }

        protected virtual T MergeFormats(List<T> formats)
        {
            return formats == null || formats.Count == 0 ? null : formats.Last();
        }
    }
}
