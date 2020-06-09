using ClosedXML.Excel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using XSheet.Renderers;
using XSheet.Renderers.ClosedXml;
using XSheet.Renderers.Formats;

namespace Xsheet.Tests
{
    public static class TestUtils
    {
        public static string GetFileName(string filename)
        {
            return $"{filename}_{DateTime.Now.ToString("HHmmss")}.xlsx";
        }

        public static void WriteDebugFile(IWorkbook wb, string filename)
        {
            var fs = File.Create(TestUtils.GetFileName(filename));
            wb.Write(fs);
        }
       
        public static string WriteDebugFile(XSheet.Renderers.Formats.IFormatApplier formatApplier, Matrix mat, string filename, IWorkbook wb = null, List<Stream> streamsToClose = null)
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

        public static string WriteDebugFile(Matrix mat, string filename, IXLWorkbook wb = null, List<Stream> streamsToClose = null)
        {
            if (wb is null)
            {
                wb = new XLWorkbook();
            }
            var rd = new ClosedXmlRenderer(wb, null);
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
