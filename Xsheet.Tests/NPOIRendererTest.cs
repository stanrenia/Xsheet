using NFluent;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using XSheetRenderers;
using Xunit;

namespace Xsheet.Tests
{
    public class NPOIRendererTest : IDisposable
    {
        private readonly IWorkbook workbook;
        private readonly IMatrixRenderer renderer;
        private readonly Dictionary<string, FileStream> fileStreamByName;
        private const string FILE_DEBUG = "debug.xlsx";
        private const string FILE_TEST_1 = "test1.xlsx";

        public NPOIRendererTest()
        {
            workbook = new XSSFWorkbook();
            renderer = new NPOIRenderer(workbook);

            fileStreamByName = new List<string> {
                FILE_DEBUG, FILE_TEST_1
            }
            .ToDictionary(name => name, name => File.Create(name));
        }

        public void Dispose()
        {
            foreach (var item in fileStreamByName)
            {
                item.Value.Close();
            }
        }

        [Fact]
        public void Should_Help_To_Debug()
        {
            var s = workbook.CreateSheet("Dumb");
            var row0 = s.CreateRow(0);
            row0.CreateCell(0).SetCellValue("[0;0]");
            row0.CreateCell(1).SetCellValue("[0;1]");
            row0.CreateCell(2).SetCellValue("[0;2]");
            var row1 = s.CreateRow(1);
            row1.CreateCell(0).SetCellValue("[1;0]");
            row1.CreateCell(1).SetCellValue("[1;1]");
            row1.CreateCell(2).SetCellValue("[1;2]");
            var row2 = s.CreateRow(2);
            row2.CreateCell(0).SetCellValue("[2;0]");
            row2.CreateCell(1).SetCellValue("[2;1]");
            row2.CreateCell(2).SetCellValue("[2;2]");
            workbook.Write(fileStreamByName[FILE_DEBUG]);
        }

        [Fact]
        public void Should_Renderer_Matrix_Basic_Example()
        {
            // GIVEN
            var cols = new List<ColumnDefinition>
            {
                new ColumnDefinition { Label = "Lastname", DataType = DataTypes.Text },
                new ColumnDefinition { Label = "Firstname", DataType = DataTypes.Text },
            };

            var values = new List<RowValue> {
                new RowValue {
                    ValuesByCol = new Dictionary<string, object> {
                        { "Lastname", "Doe" },
                        { "Firstname", "John" }
                    }
                }
            };

            var mat = Matrix.With()
                .Dimensions(2, 2)
                .Cols(cols)
                .RowValues(values)
                .Build();

            var ms = new MemoryStream();

            // WHEN
            renderer.GenerateExcelFile(mat, ms);
            ms.Close();

            WriteDebugFile(mat, FILE_TEST_1);

            // THEN
            var fileBytes = ms.ToArray();
            Check.That(fileBytes).Not.IsEmpty();

            var readWb = new XSSFWorkbook(new MemoryStream(fileBytes));
            var readSheet = readWb.GetSheetAt(0);

            var headerRow = readSheet.GetRow(0);
            Check.That(headerRow.Cells[0].StringCellValue).IsEqualTo(cols[0].Label);
            Check.That(headerRow.Cells[1].StringCellValue).IsEqualTo(cols[1].Label);

            var firstValueRow = readSheet.GetRow(1);
            Check.That(firstValueRow.Cells[0].StringCellValue).IsEqualTo(values[0].ValuesByCol["Lastname"]);
            Check.That(firstValueRow.Cells[1].StringCellValue).IsEqualTo(values[0].ValuesByCol["Firstname"]);
        }

        private void WriteDebugFile(Matrix mat, string fileName)
        {
            var wb = new XSSFWorkbook();
            var rd = new NPOIRenderer(wb);
            rd.GenerateExcelFile(mat, fileStreamByName[fileName]);
        }
    }
}
