﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Linq;
using Xsheet;

namespace XSheet.Renderers
{
    public class OpenXmlRenderer : IMatrixRenderer
    {
        private const string CHAR_FORMULA_STARTS = "=";

        public OpenXmlRenderer()
        {
        }

        public void GenerateExcelFile(Matrix mat, Stream stream)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var relationshipId = "rId1";

                //build Workbook Part
                var workbookPart = document.AddWorkbookPart();
                var workbook = new Workbook();
                var sheets = new Sheets();
                var sheet1 = new Sheet() { Name = "MySheet", SheetId = 1, Id = relationshipId };
                sheets.Append(sheet1);
                workbook.Append(sheets);
                workbookPart.Workbook = workbook;

                //build Worksheet Part
                var workSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshipId);
                var workSheet = new Worksheet();
                var sheetData = GetData(mat);
                workSheet.Append(sheetData);
                workSheetPart.Worksheet = workSheet;
            }
        }

        private SheetData GetData(Matrix mat)
        {
            var sheetData = new SheetData();
            CreateRows(mat, sheetData);
            return sheetData;
        }

        private void CreateRows(Matrix mat, SheetData sheetData)
        {
            var defaultRowDef = mat.GetRowByKey(RowDefinition.DEFAULT_KEY);
            int rowNum = mat.HasHeaders ? 1 : 0;
            foreach (var rowValue in mat.RowValues)
            {
                var rowDef = mat.GetRowByKey(rowValue.Key);
                Row row = new Row();
                foreach (var matrixCell in rowValue.Cells.OrderBy(c => c.ColIndex))
                {
                    var colDef = mat.GetOwnColumnByIndex(matrixCell.ColIndex);
                    Cell cell = new Cell();
                    //_formatApplier.ApplyFormatToCell(wb, defaultRowDef, rowDef, colDef.Index, npoiCell);
                    SetCellValue(mat, matrixCell, rowDef, colDef, cell);
                    row.Append(cell);
                }

                sheetData.Append(row);
            }
        }

        private void SetCellValue(Matrix mat, MatrixCellValue matrixCell, RowDefinition rowDef, ColumnDefinition colDef, Cell cell)
        {
            var value = matrixCell.Value;
            var dataType = colDef.DataType;

            if (matrixCell.ColName != null && rowDef.ValuesMapping.TryGetValue(matrixCell.ColName, out var func))
            {
                var calculatedValue = func.Invoke(mat, matrixCell);
                if (calculatedValue is string stringValue && stringValue.StartsWith(CHAR_FORMULA_STARTS))
                {
                    // TODO Confirm if necessary
                    //value = stringValue.Substring(1);
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
                    cell.DataType = CellValues.Number;
                    break;
                case DataTypes.Text:
                case DataTypes.Formula:
                    cell.DataType = CellValues.String;
                    break;
                default:
                    cell.DataType = CellValues.String;
                    break;
            }

            cell.CellValue = new CellValue(value.ToString());
        }
    }
}