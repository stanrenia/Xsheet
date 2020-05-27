using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xsheet;
using XSheet.Renderers.Formats;

namespace XSheet.Renderers
{
    public class NPOIRenderer : IMatrixRenderer
    {
        private readonly IWorkbook _wb;
        private readonly FormatApplier _formatApplier;

        public NPOIRenderer(IWorkbook wb, FormatApplier formatApplier)
        {
            _wb = wb;
            _formatApplier = formatApplier;
        }

        public void GenerateExcelFile(Matrix mat, Stream stream)
        {
            var sheet = _wb.CreateSheet();

            TryCreateHeaderRow(_wb, mat, sheet);

            CreateRows(_wb, mat, sheet);

            _wb.Write(stream);
        }


        private bool TryCreateHeaderRow(IWorkbook wb, Matrix mat, ISheet sheet)
        {
            if (mat.HasHeaders)
            {
                var headers = sheet.CreateRow(0);
                var cells = mat.ColumnsDefinitions.Select((colDef, colNum) =>
                {
                    var cell = headers.CreateCell(colNum);
                    _formatApplier.ApplyFormatToCell(wb, cell, colDef.HeaderCellFormat);
                    cell.SetCellValue(colDef.Label);
                    return cell;
                });

                return cells.Count() > 0;
            }

            return false;
        }

        private void CreateRows(IWorkbook wb, Matrix mat, ISheet sheet)
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
                    _formatApplier.ApplyFormatToCell(wb, defaultRowDef, rowDef, keyValue.Key, cell);
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
    }
}
