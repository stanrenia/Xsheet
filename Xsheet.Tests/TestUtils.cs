using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XSheet.Renderers.ClosedXml;
using XSheet.Renderers.NPOI;

namespace Xsheet.Tests
{
    public static class TestUtils
    {
        public static List<ICell> ReadAllCells(IWorkbook readWb, int sheetIndex)
        {
            var readSheet = readWb.GetSheetAt(sheetIndex);
            return Enumerable.Range(readSheet.FirstRowNum, readSheet.LastRowNum)
                .SelectMany(i =>
                {
                    var curRow = readSheet.GetRow(i);
                    return Enumerable.Range(curRow.FirstCellNum, curRow.LastCellNum)
                        .Select(j => curRow.GetCell(j));
                })
                .ToList();
        }

        public static string GetFileName(string filename)
        {
            return $"{filename}_{DateTime.Now.ToString("HHmmss")}.xlsx";
        }

        public static void WriteDebugFile(IWorkbook wb, string filename)
        {
            var fs = File.Create(TestUtils.GetFileName(filename));
            wb.Write(fs);
        }
       
        public static string WriteDebugFile(XSheet.Renderers.NPOI.Formats.IFormatApplier formatApplier, Matrix mat, string filename, IWorkbook wb = null, List<Stream> streamsToClose = null)
        {
            if (wb is null)
            {
                wb = new XSSFWorkbook();
            }
            var rd = new NPOIRenderer(wb, formatApplier);
            var targetFilename = TestUtils.GetFileName(filename);
            var fs = File.Create(targetFilename);
            
            if (streamsToClose != null) { 
                streamsToClose.Add(fs);
            }
            
            rd.GenerateExcelFile(mat, fs);
            return targetFilename;
        }

        public static string WriteDebugFile(Matrix mat, string filename, IXLWorkbook wb = null, List<Stream> streamsToClose = null, IFormatApplier formatApplier = null)
        {
            if (wb is null)
            {
                wb = new XLWorkbook();
            }
            if (formatApplier is null)
            {
                formatApplier = new ClosedXmlFormatApplier();
            }

            var rd = new ClosedXmlRenderer(wb, formatApplier);
            var targetFilename = TestUtils.GetFileName(filename);
            var fs = File.Create(targetFilename);
            if (streamsToClose != null)
            {
                streamsToClose.Add(fs);
            }
            rd.GenerateExcelFile(mat, fs);
            return targetFilename;
        }
    }
}
