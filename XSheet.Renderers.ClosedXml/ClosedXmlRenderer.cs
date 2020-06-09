using ClosedXML.Excel;
using System;
using System.IO;
using System.Linq;
using Xsheet;

namespace XSheet.Renderers.ClosedXml
{
    public class ClosedXmlRenderer : IMatrixRenderer
    {
        private readonly IXLWorkbook _wb;
        private readonly IFormatApplier _formatApplier;

        public ClosedXmlRenderer(IXLWorkbook wb, IFormatApplier formatApplier)
        {
            _wb = wb;
            _formatApplier = formatApplier;
        }

        public void GenerateExcelFile(Matrix mat, Stream stream)
        {
            var sheet = _wb.Worksheets.Add("Sheet 1");

            TryCreateHeaderRow(_wb, mat, sheet);

            CreateRows(_wb, mat, sheet);

            _wb.SaveAs(stream);
        }

        private void TryCreateHeaderRow(IXLWorkbook wb, Matrix mat, IXLWorksheet sheet)
        {
            if (mat.HasHeaders)
            {
                var headers = sheet.Row(1);
                mat.ColumnsDefinitions.Select((colDef, colNum) =>
                {
                    var cell = headers.Cell(1 + colDef.Index);
                    //_formatApplier.ApplyFormatToCell(wb, cell, colDef.HeaderCellFormat);
                    cell.Value = colDef.Label;
                    return cell;
                }).ToArray();
            }
        }

        private void CreateRows(IXLWorkbook wb, Matrix mat, IXLWorksheet sheet)
        {
            var defaultRowDef = mat.GetRowByKey(RowDefinition.DEFAULT_KEY);
            int rowNum = mat.HasHeaders ? 1 : 0;
            foreach (var rowValue in mat.RowValues)
            {
                var rowDef = mat.GetRowByKey(rowValue.Key);
                var row = sheet.Row(1 + rowNum++);
                foreach (var matrixCell in rowValue.Cells.OrderBy(c => c.ColIndex))
                {
                    var colDef = mat.GetOwnColumnByIndex(matrixCell.ColIndex);
                    var cell = row.Cell(1 + colDef.Index);
                    //_formatApplier.ApplyFormatToCell(wb, defaultRowDef, rowDef, colDef.Index, cell);
                    SetCellValue(mat, matrixCell, rowDef, colDef, cell);
                }
            }
        }

        private void SetCellValue(Matrix mat, MatrixCellValue matrixCell, RowDefinition rowDef, ColumnDefinition colDef, IXLCell cell)
        {
            var value = matrixCell.Value;
            bool isFormula = false;

            if (matrixCell.ColName != null && rowDef.ValuesMapping.TryGetValue(matrixCell.ColName, out var func))
            {
                var calculatedValue = func.Invoke(mat, matrixCell);
                if (calculatedValue is string stringValue && stringValue.StartsWith(MatrixConstants.CHAR_FORMULA_STARTS))
                {
                    value = stringValue.Substring(1);
                    isFormula = true;
                }
                else
                {
                    value = calculatedValue;
                }
            }

            if (value is null)
            {
                return;
            }

            if (isFormula)
            {
                cell.FormulaA1 = Convert.ToString(value);
            }
            else
            {
                cell.Value = value;
            }
        }
    }
}
