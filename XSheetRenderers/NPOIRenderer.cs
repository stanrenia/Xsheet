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

            TryCreateHeaderRow(mat, sheet);
            
            CreateRows(mat, sheet);

            _wb.Write(stream);
        }


        private static bool TryCreateHeaderRow(Matrix mat, ISheet sheet)
        {
            if (mat.HasHeaders)
            {
                var headers = sheet.CreateRow(0);
                var cells = mat.ColumnsDefinitions.Select((col, colNum) =>
                {
                    var cell = headers.CreateCell(colNum);
                    cell.SetCellValue(col.Label);
                    return cell;
                });

                return cells.Count() > 0;
            }

            return false;
        }
        
        private static void CreateRows(Matrix mat, ISheet sheet)
        {
            int rowNum = mat.HasHeaders ? 1 : 0;
            foreach (var rowValue in mat.RowValues)
            {
                var row = sheet.CreateRow(rowNum++);
                foreach (var keyValue in rowValue.ValuesByCol)
                {
                    var col = mat.GetColumnByKey(keyValue.Key);
                    var cell = row.CreateCell(col.Index);
                    var cellValue = ConvertValue(keyValue, col);
                    cell.SetCellValue(cellValue);
                }
            }
        }

        private static string ConvertValue(KeyValuePair<string, object> keyValue, ColumnDefinition col)
        {
            switch (col.DataType)
            {
                case DataTypes.Text:
                    return Convert.ToString(keyValue.Value);
                default:
                    return Convert.ToString(keyValue.Value);
            }
        }
    }
}
