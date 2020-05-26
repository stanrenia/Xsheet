using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xsheet;

namespace XSheetRenderers
{
    public class NPOIRenderer : IMatrixRenderer
    {
        private readonly IWorkbook _wb;

        public NPOIRenderer(IWorkbook wb)
        {
            _wb = wb;
        }

        public void GenerateExcelFile(Matrix mat, Stream stream)
        {
            var sheet = _wb.CreateSheet();

            TryCreateHeaderRow(_wb, mat, sheet);
            
            CreateRows(_wb, mat, sheet);

            _wb.Write(stream);
        }


        private static bool TryCreateHeaderRow(IWorkbook wb, Matrix mat, ISheet sheet)
        {
            if (mat.HasHeaders)
            {
                var headers = sheet.CreateRow(0);
                var cells = mat.ColumnsDefinitions.Select((colDef, colNum) =>
                {
                    var cell = headers.CreateCell(colNum);
                    ApplyFormatToCell(wb, cell, colDef.HeaderCellFormat);
                    cell.SetCellValue(colDef.Label);
                    return cell;
                });

                return cells.Count() > 0;
            }

            return false;
        }

        private static void CreateRows(IWorkbook wb, Matrix mat, ISheet sheet)
        {
            var defaultRowDef = mat.GetRowByKey(RowDefinition.DEFAULT_KEY);
            int rowNum = mat.HasHeaders ? 1 : 0;
            foreach (var rowValue in mat.RowValues)
            {
                var rowDef = mat.GetRowByKey(rowValue.Key);
                var row = sheet.CreateRow(rowNum++);
                foreach (var keyValue in rowValue.ValuesByCol)
                {
                    var colDef = mat.GetColumnByKey(keyValue.Key);
                    var cell = row.CreateCell(colDef.Index);
                    ApplyFormatToCell(wb, defaultRowDef, rowDef, keyValue.Key, cell);
                    SetCellValue(keyValue, colDef, cell);
                }
            }
        }

        private static void SetCellValue(KeyValuePair<string, object> keyValue, ColumnDefinition col, ICell cell)
        {
            switch (col.DataType)
            {
                case DataTypes.Number:
                    cell.SetCellValue(Convert.ToDouble(keyValue.Value));
                    break;
                case DataTypes.Text:
                    cell.SetCellValue(Convert.ToString(keyValue.Value));
                    break;
                default:
                    cell.SetCellValue(Convert.ToString(keyValue.Value));
                    break;
            }
        }

        // Apply default format of default row 
        // Apply each Col format of default row
        // Apply default format of current row
        // Apply each Col format of current row
        private static void ApplyFormatToCell(IWorkbook wb, RowDefinition defaultRowDef, RowDefinition rowDef, string columnKey, ICell cell)
        {
            var formats = new List<Format>();
            void AddFormatIfNotNull(Format f)
            {
                if (f != null)
                {
                    formats.Add(f);
                }
            }

            AddFormatIfNotNull(defaultRowDef.DefaultCellFormat);
            defaultRowDef.FormatsByCol.TryGetValue(columnKey, out Format defaultColFormat);
            AddFormatIfNotNull(defaultColFormat);
            AddFormatIfNotNull(rowDef.DefaultCellFormat);
            rowDef.FormatsByCol.TryGetValue(columnKey, out Format colFormat);
            AddFormatIfNotNull(colFormat);
            
            var mergedFormat = Format.MergeFormats(formats);
            
            ApplyFormatToCell(wb, cell, mergedFormat);
        }

        private static void ApplyFormatToCell(IWorkbook wb, ICell cell, Format format)
        {
            if (format is null)
            {
                return;
            }

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
    }
}
