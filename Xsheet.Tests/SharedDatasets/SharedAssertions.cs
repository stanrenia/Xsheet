using NFluent;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Linq;

namespace Xsheet.Tests.SharedDatasets
{
    /// <summary>
    /// Shared assertions using NPOI to read Excel files
    /// </summary>
    public static class SharedAssertions
    {
        public static void Assert_Basic_Example(MemoryStream ms, MatrixDataset dataset)
        {
            var fileBytes = ms.ToArray();
            Check.That(fileBytes).Not.IsEmpty();

            var readWb = new XSSFWorkbook(new MemoryStream(fileBytes));
            var readSheet = readWb.GetSheetAt(0);

            var cols = dataset.Cols;
            var values = dataset.Values;

            var headerRow = readSheet.GetRow(0);
            Check.That(headerRow.Cells[0].StringCellValue).IsEqualTo(cols[0].Label);
            Check.That(headerRow.Cells[1].StringCellValue).IsEqualTo(cols[1].Label);

            var firstValueRow = readSheet.GetRow(1);
            Check.That(firstValueRow.Cells[0].StringCellValue).IsEqualTo(values[0].ValuesByColName["Lastname"]);
            Check.That(firstValueRow.Cells[1].StringCellValue).IsEqualTo(values[0].ValuesByColName["Firstname"]);
        }

        public static void Assert_Matrix_Cells_Lookup_On_A_Single_Matrix(string filename)
        {
            var readWb = new XSSFWorkbook(File.OpenRead(filename));
            var readSheet = readWb.GetSheetAt(0);

            var row1 = readSheet.GetRow(1);
            Check.That(row1.Cells[0].StringCellValue).IsEqualTo("Mario");
            Check.That(row1.Cells.Last().CellFormula).IsEqualTo("AVERAGE(10,20,30)");
            Check.That(row1.Cells.Skip(1).Extracting("NumericCellValue")).ContainsExactly(10, 20, 30, 60, 0);

            var row2 = readSheet.GetRow(2);
            Check.That(row2.Cells[0].StringCellValue).IsEqualTo("Luigi");
            Check.That(row2.Cells.Last().CellFormula).IsEqualTo("AVERAGE(12,23,34)");
            Check.That(row2.Cells.Skip(1).Extracting("NumericCellValue")).ContainsExactly(12, 23, 34, 69, 0);

            var row3 = readSheet.GetRow(3);
            Check.That(row3.Cells[0].StringCellValue).IsEqualTo("TOTAL");
            Check.That(row3.Cells.Skip(1).Extracting("NumericCellValue")).ContainsExactly(22, 0, 0, 0, 0);
            Check.That(row3.Cells[2].CellFormula).IsEqualTo("SUM(C2:C3)");
            Check.That(row3.Cells[3].CellFormula).IsEqualTo("D2+D3");
            Check.That(row3.Cells[4].CellFormula).IsEqualTo("SUM(E2:E3)");
            Check.That(row3.Cells[5].CellFormula).IsEqualTo("AVERAGE(F2:F3)");
        }

        public static void Assert_Format_With_Specific_FormatApplier(MemoryStream ms, FormatDataset dataset)
        {
            var fileBytes = ms.ToArray();
            Check.That(fileBytes).Not.IsEmpty();

            var readWb = new XSSFWorkbook(new MemoryStream(fileBytes));
            var readSheet = readWb.GetSheetAt(0);

            // Headers
            var headerRow = readSheet.GetRow(0);
            // -- Firstname cell is Italic
            Check.That(headerRow.Cells[1].CellStyle.GetFont(readWb).IsItalic).IsTrue();

            // Row 1
            var rowValue1 = readSheet.GetRow(1);
            var lastnameFont1 = rowValue1.Cells[0].CellStyle.GetFont(readWb);
            // -- Lastname is Bold
            Check.That(lastnameFont1.IsBold).IsTrue();
            // -- Age has BgColor=LightGrey
            Check.That(rowValue1.Cells[2].CellStyle.FillForegroundColorColor.ToARGB()).IsEqualTo(dataset.ColorLightGrey);

            // Row 2
            var rowValue2 = readSheet.GetRow(2);
            // -- Age has BgColor=Blue
            Check.That(rowValue2.Cells[2].CellStyle.FillForegroundColorColor.ToARGB()).IsEqualTo(dataset.ColorBlue);
            // -- Lastname is Bold
            Check.That(rowValue2.Cells[0].CellStyle.GetFont(readWb).IsBold).IsTrue();
            // -- Firstname is not Bold
            Check.That(rowValue2.Cells[1].CellStyle.GetFont(readWb).IsBold).IsFalse();

            // Row 3
            var rowValue3 = readSheet.GetRow(3);
            // -- Lastname is Bold
            Check.That(rowValue3.Cells[0].CellStyle.GetFont(readWb).IsBold).IsTrue();
            // -- Age has BgColor=LightGrey
            Check.That(rowValue3.Cells[2].CellStyle.FillForegroundColorColor.ToARGB()).IsEqualTo(dataset.ColorLightGrey);

            // All cells (skip headers)
            // -- FontSize is 12
            Check.That(TestUtils.ReadAllCells(readWb, 0).Skip(dataset.ColsCount)).ContainsOnlyElementsThatMatch(cell => cell.CellStyle.GetFont(readWb).FontHeightInPoints == 12);
        }
    }
}
