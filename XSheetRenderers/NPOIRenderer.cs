﻿using NPOI.SS.UserModel;
using System;
using System.IO;
using System.Linq;
using Xsheet;
using XSheet.Renderers.Formats;

namespace XSheet.Renderers
{
    public class NPOIRenderer : IMatrixRenderer
    {
        private const string CHAR_FORMULA_STARTS = "=";
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
                foreach (var matrixCell in rowValue.Cells.OrderBy(c => c.ColIndex))
                {
                    var colDef = mat.GetOwnColumnByIndex(matrixCell.ColIndex);
                    var npoiCell = row.CreateCell(colDef.Index);
                    _formatApplier.ApplyFormatToCell(wb, defaultRowDef, rowDef, colDef.Index, npoiCell);
                    SetCellValue(mat, matrixCell, rowDef, colDef, npoiCell);
                }
            }
        }

        private static void SetCellValue(Matrix mat, MatrixCellValue matrixCell, RowDefinition rowDef, ColumnDefinition colDef, ICell npoiCell)
        {
            var value = matrixCell.Value;
            var dataType = colDef.DataType;

            if (matrixCell.ColName != null && rowDef.ValuesMapping.TryGetValue(matrixCell.ColName, out var func))
            {
                var calculatedValue = func.Invoke(mat, matrixCell);
                if (calculatedValue is string stringValue && stringValue.StartsWith(CHAR_FORMULA_STARTS))
                {
                    value = stringValue.Substring(1);
                    dataType = DataTypes.Formula;
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

            switch (dataType)
            {
                case DataTypes.Number:
                    npoiCell.SetCellValue(Convert.ToDouble(value));
                    break;
                case DataTypes.Text:
                    npoiCell.SetCellValue(Convert.ToString(value));
                    break;
                case DataTypes.Formula:
                    npoiCell.SetCellFormula(Convert.ToString(value));
                    break;
                default:
                    npoiCell.SetCellValue(Convert.ToString(value));
                    break;
            }
        }
    }
}
