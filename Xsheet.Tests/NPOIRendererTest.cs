using NFluent;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using XSheet.Renderers;
using XSheet.Renderers.Formats;
using Xunit;

namespace Xsheet.Tests
{
    public class NPOIRendererTest : IDisposable
    {
        private readonly IWorkbook _workbook;
        private readonly FormatApplier _defaultFormatApplier;
        private readonly IMatrixRenderer _renderer;
        private readonly List<Stream> _fileStreamToClose;
        private const string FILE_DEBUG = "debug";
        private const string FILE_DEBUG2 = "debug2";
        private const string FILE_TEST_1 = "test1";
        private const string FILE_TEST_FORMAT_BASIC = "test_format_basic";
        private const string FILE_TEST_FORMAT_NPOI = "test_format_npoi";

        public NPOIRendererTest()
        {
            _workbook = new XSSFWorkbook();
            _defaultFormatApplier = new NPOIFormatApplier();
            _renderer = new NPOIRenderer(_workbook, _defaultFormatApplier);
            _fileStreamToClose = new List<Stream>();
        }

        public void Dispose()
        {
            foreach (var stream in _fileStreamToClose)
            {
                stream.Close();
            }
        }

        [Fact]
        public void Should_Help_To_Debug()
        {
            var s = _workbook.CreateSheet("Dumb");
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
            var fs = File.Create(FILE_DEBUG);
            _fileStreamToClose.Add(fs);
            _workbook.Write(fs);
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
            _renderer.GenerateExcelFile(mat, ms);
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

        [Fact]
        public void Should_Renderer_Matrix_With_Basic_Format()
        {
            // GIVEN
            const string Lastname = "Lastname";
            const string Firstname = "Firstname";
            const string Age = "Age";
            var cols = new List<ColumnDefinition>
            {
                new ColumnDefinition { Label = Lastname, DataType = DataTypes.Text },
                new ColumnDefinition { Label = Firstname, DataType = DataTypes.Text, HeaderCellFormat = new BasicFormat { FontStyle = FontStyle.Italic } },
                new ColumnDefinition { Label = Age, DataType = DataTypes.Number },
            };

            const string Even = "EVEN";
            const string Odd = "ODD";

            var ColorBlueIndex = IndexedColors.LightBlue.Index;
            var ColorLightGreyIndex = IndexedColors.Grey25Percent.Index;
            var rows = new List<RowDefinition>
            {
                new RowDefinition { 
                    DefaultCellFormat = new BasicFormat { FontSize = 12 }, 
                    FormatsByCol = new Dictionary<string, Format> {
                        { Lastname, new BasicFormat { FontStyle = FontStyle.Bold } },
                        { Age, new BasicFormat { BackgroundColor = ColorLightGreyIndex.ToString() } }
                    }
                },
                new RowDefinition { 
                    Key = Odd, 
                    FormatsByCol = new Dictionary<string, Format> {
                        { Age, new BasicFormat { BackgroundColor = ColorBlueIndex.ToString() } }
                    } 
                },
            };

            var values = new List<RowValue> {
                new RowValue {
                    Key = Even,
                    ValuesByCol = new Dictionary<string, object> {
                        { Lastname, "Doe" },
                        { Firstname, "John" },
                        { Age, 30 }
                    }
                },
                new RowValue {
                    Key = Odd,
                    ValuesByCol = new Dictionary<string, object> {
                        { Lastname, "Clinton" },
                        { Firstname, "Bob" },
                        { Age, 41 }
                    }
                },
                new RowValue {
                    Key = Even,
                    ValuesByCol = new Dictionary<string, object> {
                        { Lastname, "Doa" },
                        { Firstname, "Johana" },
                        { Age, 36 }
                    }
                }
            };

            var mat = Matrix.With()
                .Cols(cols)
                .Rows(rows)
                .RowValues(values)
                .Build();

            var ms = new MemoryStream();

            var formatApplier = new BasicFormatApplier();
            var renderer = new NPOIRenderer(_workbook, formatApplier);

            // WHEN
            renderer.GenerateExcelFile(mat, ms);
            ms.Close();

            WriteDebugFile(formatApplier, mat, FILE_TEST_FORMAT_BASIC);

            // THEN
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
            Check.That(rowValue1.Cells[2].CellStyle.FillForegroundColor).IsEqualTo(ColorLightGreyIndex);

            // Row 2
            var rowValue2 = readSheet.GetRow(2);
            // -- Age has BgColor=Blue
            Check.That(rowValue2.Cells[2].CellStyle.FillForegroundColor).IsEqualTo(ColorBlueIndex);
            // -- Lastname is Bold
            Check.That(rowValue2.Cells[0].CellStyle.GetFont(readWb).IsBold).IsTrue();

            // Row 3
            var rowValue3 = readSheet.GetRow(3);
            // -- Lastname is Bold
            Check.That(rowValue3.Cells[0].CellStyle.GetFont(readWb).IsBold).IsTrue();
            // -- Age has BgColor=LightGrey
            Check.That(rowValue3.Cells[2].CellStyle.FillForegroundColor).IsEqualTo(ColorLightGreyIndex);

            // All cells (skip headers)
            // -- FontSize is 12
            Check.That(ReadAllCells(readWb, 0).Skip(cols.Count)).ContainsOnlyElementsThatMatch(cell => cell.CellStyle.GetFont(readWb).FontHeightInPoints == 12);
        }

        [Fact]
        public void Should_Renderer_Matrix_With_NPOIFormat()
        {
            // GIVEN
            const string Lastname = "Lastname";
            const string Firstname = "Firstname";
            const string Age = "Age";

            IFont Font1()
            {
                var f = _workbook.CreateFont();
                f.IsItalic = true;
                return f;
            }

            IFont Font2()
            {
                var f = _workbook.CreateFont();
                f.FontHeightInPoints = 12;
                return f;
            }

            IFont Font3()
            {
                var f = Font2();
                f.IsBold = true;
                return f;
            }

            var style1 = _workbook.CreateCellStyle();
            style1.SetFont(Font1());

            var style2 = _workbook.CreateCellStyle();
            style2.SetFont(Font2());

            var style3 = _workbook.CreateCellStyle();
            style3.CloneStyleFrom(style2);
            style3.SetFont(Font3());

            var ColorBlueIndex = IndexedColors.LightBlue.Index;
            var ColorLightGreyIndex = IndexedColors.Grey25Percent.Index;
            
            var style4 = _workbook.CreateCellStyle();
            style4.CloneStyleFrom(style2);
            style4.FillPattern = FillPattern.SolidForeground;
            style4.FillForegroundColor = ColorLightGreyIndex;

            var style5 = _workbook.CreateCellStyle();
            style5.CloneStyleFrom(style4);
            style5.FillForegroundColor = ColorBlueIndex;

            var cols = new List<ColumnDefinition>
            {
                new ColumnDefinition { Label = Lastname, DataType = DataTypes.Text },
                new ColumnDefinition { Label = Firstname, DataType = DataTypes.Text, HeaderCellFormat = new NPOIFormat { CellStyle = style1 } },
                new ColumnDefinition { Label = Age, DataType = DataTypes.Number },
            };

            const string Even = "EVEN";
            const string Odd = "ODD";

            var rows = new List<RowDefinition>
            {
                new RowDefinition {
                    DefaultCellFormat = new NPOIFormat { CellStyle = style2 },
                    FormatsByCol = new Dictionary<string, Format> {
                        { Lastname, new NPOIFormat { CellStyle = style3 } },
                        { Age, new NPOIFormat { CellStyle = style4 } }
                    }
                },
                new RowDefinition {
                    Key = Odd,
                    FormatsByCol = new Dictionary<string, Format> {
                        { Age, new NPOIFormat { CellStyle = style5 } }
                    }
                },
            };

            var values = new List<RowValue> {
                new RowValue {
                    Key = Even,
                    ValuesByCol = new Dictionary<string, object> {
                        { Lastname, "Doe" },
                        { Firstname, "John" },
                        { Age, 30 }
                    }
                },
                new RowValue {
                    Key = Odd,
                    ValuesByCol = new Dictionary<string, object> {
                        { Lastname, "Clinton" },
                        { Firstname, "Bob" },
                        { Age, 41 }
                    }
                },
                new RowValue {
                    Key = Even,
                    ValuesByCol = new Dictionary<string, object> {
                        { Lastname, "Doa" },
                        { Firstname, "Johana" },
                        { Age, 36 }
                    }
                }
            };

            var mat = Matrix.With()
                .Cols(cols)
                .Rows(rows)
                .RowValues(values)
                .Build();

            var ms = new MemoryStream();

            // WHEN
            _renderer.GenerateExcelFile(mat, ms);
            ms.Close();

            WriteDebugFile(_defaultFormatApplier, mat, FILE_TEST_FORMAT_NPOI, _workbook);

            // THEN
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
            Check.That(rowValue1.Cells[2].CellStyle.FillForegroundColor).IsEqualTo(ColorLightGreyIndex);

            // Row 2
            var rowValue2 = readSheet.GetRow(2);
            // -- Age has BgColor=Blue
            Check.That(rowValue2.Cells[2].CellStyle.FillForegroundColor).IsEqualTo(ColorBlueIndex);
            // -- Lastname is Bold
            Check.That(rowValue2.Cells[0].CellStyle.GetFont(readWb).IsBold).IsTrue();
            // -- Firstname is not Bold
            Check.That(rowValue2.Cells[1].CellStyle.GetFont(readWb).IsBold).IsFalse();

            // Row 3
            var rowValue3 = readSheet.GetRow(3);
            // -- Lastname is Bold
            Check.That(rowValue3.Cells[0].CellStyle.GetFont(readWb).IsBold).IsTrue();
            // -- Age has BgColor=LightGrey
            Check.That(rowValue3.Cells[2].CellStyle.FillForegroundColor).IsEqualTo(ColorLightGreyIndex);

            // All cells (skip headers)
            // -- FontSize is 12
            Check.That(ReadAllCells(readWb, 0).Skip(cols.Count)).ContainsOnlyElementsThatMatch(cell => cell.CellStyle.GetFont(readWb).FontHeightInPoints == 12);
        }

        [Fact]
        public void Should_Help_To_Debug_NPOIFormat()
        {
            var wb = new XSSFWorkbook();
            
            short colorBlack = IndexedColors.Black.Index;
            short colorLightBlue = IndexedColors.LightBlue.Index;
            short colorLightOrange = IndexedColors.LightOrange.Index;
            short colorGreen = IndexedColors.Green.Index;

            var sheet = wb.CreateSheet();
            var row0 = sheet.CreateRow(0);
            var cell0 = row0.CreateCell(0);
            cell0.SetCellValue("FIRST");
            var cell1 = row0.CreateCell(1);
            cell1.SetCellValue("SECOND");

            // Best way to clone Font: Wrap its creation in a method
            // Because the existing method IFont.CloneStyleFrom(IFont) is buggy
            IFont GetFont1()
            {
                var f = wb.CreateFont();
                f.Color = colorBlack;
                f.FontHeightInPoints = 14;
                f.IsBold = true;
                return f;
            }

            var s1 = wb.CreateCellStyle();
            s1.FillPattern = FillPattern.SolidForeground;
            s1.FillForegroundColor = colorLightBlue;
            s1.VerticalAlignment = VerticalAlignment.Bottom;
            var f1 = GetFont1();
            s1.SetFont(f1);
            cell0.CellStyle = s1;

            var s2 = wb.CreateCellStyle();
            s2.CloneStyleFrom(s1);
            s2.FillForegroundColor = colorLightOrange;
            s2.Rotation = 49;
            s1.VerticalAlignment = VerticalAlignment.Top;
            var f2 = GetFont1();
            f2.IsBold = false;
            f2.IsItalic = true;
            f2.Color = colorGreen;
            s2.SetFont(f2);
            cell1.CellStyle = s2;

            WriteDebugFile(wb, FILE_DEBUG2);

            using (var ms = new MemoryStream())
            {
                wb.Write(ms);

                var readWb = new XSSFWorkbook(new MemoryStream(ms.ToArray()));
                var readSheet = readWb.GetSheetAt(0);
                var rs1 = readSheet.GetRow(0).GetCell(0).CellStyle;
                Check.That(rs1).IsNotNull();
                Check.That(rs1.FillForegroundColor).IsEqualTo(colorLightBlue);
                Check.That(rs1.GetFont(readWb).IsBold).IsTrue();
                Check.That(rs1.GetFont(readWb).IsItalic).IsFalse();
                Check.That(rs1.GetFont(readWb).Color).IsEqualTo(colorBlack);
                var rs2 = readSheet.GetRow(0).GetCell(1).CellStyle;
                Check.That(rs2.GetFont(readWb).IsBold).IsFalse();
                Check.That(rs2.GetFont(readWb).IsItalic).IsTrue();
                Check.That(rs2.GetFont(readWb).Color).IsEqualTo(colorGreen);
            }
        }

        private List<ICell> ReadAllCells(IWorkbook readWb, int sheetIndex)
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

        private void WriteDebugFile(Matrix mat, string fileName)
        {
            WriteDebugFile(_defaultFormatApplier, mat, fileName);
        }

        private void WriteDebugFile(FormatApplier formatApplier, Matrix mat, string filename, IWorkbook wb = null)
        {
            if (wb is null)
            {
                wb = new XSSFWorkbook();
            }
            var rd = new NPOIRenderer(wb, formatApplier);
            var fs = File.Create($"{filename}_{DateTime.Now.ToString("HHmmss")}.xlsx");
            _fileStreamToClose.Add(fs);
            rd.GenerateExcelFile(mat, fs);
        }

        private void WriteDebugFile(IWorkbook wb, string filename)
        {
            var fs = File.Create($"{filename}_{DateTime.Now.ToString("HHmmss")}.xlsx");
            wb.Write(fs);
        }

        private void WriteDebugFileAndClose(IWorkbook wb, string filename)
        {
            var fs = File.Create($"{filename}_{DateTime.Now.ToString("HHmmss")}.xlsx");
            wb.Write(fs);
            fs.Close();
            wb.Close();
        }
    }
}
