using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using Xsheet;
using Xsheet.Formats;

namespace XSheet.Renderers.NPOI.Formats
{
    public abstract class BaseFormatApplier<T> : IFormatApplier
        where T : class, IFormat
    {
        /// <summary>
        /// Get All Formats ordered by less prioritary for the given Cell, then call MergeFormats()
        /// Override MergeFormats() to change the default behavior, which is only taking the 4th format (the last one)
        /// </summary>
        public void ApplyFormatToCell(IWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, int columnIndex, ICell cell)
        {
            List<T> formats = FormatMerger.GetFormatsOrderedByLessPrioritary<T>(defaultRowDef, rowDef, columnIndex);
            var mergedFormat = MergeFormats(formats);

            if (mergedFormat != null)
            {
                ApplyFormatToCell(wb, cell, mergedFormat);
            }
        }

        public abstract void ApplyFormatToCell(IWorkbook wb, ICell cell, T format);

        public void ApplyFormatToCell(IWorkbook wb, ICell cell, IFormat format)
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
