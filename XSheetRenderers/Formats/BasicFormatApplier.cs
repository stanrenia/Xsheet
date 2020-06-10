using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using Xsheet;

namespace XSheet.Renderers.NPOI.Formats
{
    public class BasicFormatApplier : BaseFormatApplier<BasicFormat>
    {
        public override void ApplyFormatToCell(IWorkbook wb, ICell cell, BasicFormat format)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            IFont font = wb.CreateFont();

            if (!string.IsNullOrEmpty(format.BackgroundColor))
            {
                // NPOI: FillForegroundColor refers to the "Fill Color" on Excel
                cellStyle.FillForegroundColor = Convert.ToInt16(format.BackgroundColor);
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }

            if (format.FontSize.HasValue)
            {
                font.FontHeightInPoints = format.FontSize.Value;
            }

            if (!string.IsNullOrEmpty(format.FontStyle))
            {
                font.IsBold = format.FontStyle == FontStyle.Bold;
                font.IsItalic = format.FontStyle == FontStyle.Italic;
                font.IsStrikeout = format.FontStyle == FontStyle.Strikeout;
            }

            cellStyle.SetFont(font);
            cell.CellStyle = cellStyle;
        }

        protected override BasicFormat MergeFormats(List<BasicFormat> formats)
        {
            return BasicFormat.MergeFormats(formats);
        }
    }
}
