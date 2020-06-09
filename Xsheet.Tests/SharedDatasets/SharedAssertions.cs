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
    }
}
