# Xsheet - High level C# API to build Excel report easily
**Compatible with :**
- NPOI
- ClosedXML
- OpenXML (only values are supported)
  - OpenXML is too verbose to be implemented quickly, ClosedXML is the preferred approach for the moment.

## See various examples at ```Xsheet.Tests/Renderers/*.cs```
### Advanced example
![Advanced example screenshot](https://github.com/stanrenia/Xsheet/blob/master/Screenshots/capt2.PNG)

```csharp
var mat = Matrix.With()
// Columns definitions
.Cols()
    // ColumnDefinition: Defines Column Name, Label (displayed text) and cells values types
    // Name property is used as a reference in RowDefinition and RowValue classes
    .Col(Playername, "Player name")
    .Col(Score1, "Score 1", DataTypes.Number)
    .Col(Score2, "Score 2", DataTypes.Number)
    .Col(Score3, "Score 3", DataTypes.Number)
    .Col(Total, "Total", DataTypes.Number)
    .Col(Mean, "Mean", DataTypes.Number)
// Rows definitions
.Rows()
    // RowDefinition: Defines how to render cells (their values and styles) for RowValues having the same Key as the row definition.
    // Below, the RowDefinition is the default since 'Key' is not defined (null)
    // By default, each RowValue will get this RowDefinition
    .Row()
        // ValuesMapping property lets you define Cells values by accessing the whole Matrix
        // For this row, we define Total and Mean columns
        // .ValueMap() is a shortcut method reducing the boilerplate of Dictionnary instanciation
        
        // 1) Compute the SUM at runtime using C#
        // As Value is of type object, we have to convert it.
        // Note: this is only for demonstration purpose, it's preferable to compute it before
        .ValueMap(Total, (mat, cell) =>
        {
            var row = mat.Row(cell);
            return Convert.ToDouble(row.Col(Score1).Value)
            + Convert.ToDouble(row.Col(Score2).Value)
            + Convert.ToDouble(row.Col(Score3).Value);
        })
        // 2) Define a formula using Cells values. Eg: =AVERAGE(10,20,30)
        .ValueMap(Mean, (mat, cell) =>
        {
            var row = mat.Row(cell);
            return $"=AVERAGE({row.Col(Score1).Value},{row.Col(Score2).Value},{row.Col(Score3).Value})";
        })
    // We add another RowDefinition representing the final row (at the bottom)
    // It's a summary row, so we're doing SUM and AVERAGE operations
    // For demonstration purpose, we write theses operations in various ways
    // Below, we set the Key to 'FinalTotal' since it's not the default RowDefinition
    // Also we set some styling with 'format'
    .Row(FinalTotal, format)
        // 3) Set the value to "TOTAL"
        .ValueMap(Playername, (mat, cell) => "TOTAL")
        // 4) Compute the SUM of column Score1, at runtime
        .ValueMap(Score1, (mat, cell) => mat.Col(cell).Values.Select(v => Convert.ToDouble(v)).Sum())
        // 5) Define the SUM of column Score2 using Excel's SUM formula with Cells Addresses
        // FROM the top 1st row TO the row above the current 'FinalTotal' row
        .ValueMap(Score2, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Score2).Address})")
        // 6) Excel's custom formula Eg: =D2+D3+D4
        .ValueMap(Score3, (mat, cell) => $"={mat.Col(cell).Cells.SkipLast(1).Addresses("+")}")
        // 7) Same as 5)
        .ValueMap(Total, (mat, cell) => $"=SUM({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Total).Address})")
        // 8) Using AVERAGE formula
        .ValueMap(Mean, (mat, cell) => $"=AVERAGE({mat.Col(cell).Cells[0].Address}:{mat.RowAbove(cell).Col(Mean).Address})")

// Here we passe the values for each row
// For demonstration purpose, we create the values right here
// However, it's preferable to transform your business data to a list of RowValue
// before, as it could be complex.
.RowValues(new List<RowValue> {
    // Mario and Luigi rows will get the default RowDefinition
    new RowValue { ValuesByColName = new Dictionary<string, object> {
        { Playername, "Mario" }, { Score1, 10 }, { Score2, 20 }, { Score3, 30 }
    } },
    new RowValue { ValuesByColName = new Dictionary<string, object> {
        { Playername, "Luigi" }, { Score1, 12 }, { Score2, 23 }, { Score3, 34 }
    } },
    // This RowValue is a special one as it does not contain Data, it only defines the row's position (at the bottom)
    // It will get the FinalTotal RowDefinition
    new RowValue { Key = FinalTotal }
})
.Build();

// Generate the Excel file - using NPOI
var wb = new XSSFWorkbook();

using (var fs = File.Create("report.xlsx")) 
{
    var rd = new NPOIRenderer(wb, new NPOIFormatApplier());
    rd.GenerateExcelFile(mat, fs);
}
```

### Styling
With ```NPOIFormat```
![Styling with NPOIFormat screenshot](https://github.com/stanrenia/Xsheet/blob/master/Screenshots/capt1.PNG)

```csharp
var _workbook = new XSSFWorkbook();

// Use your own methods in order to share Fonts
// Because NPOI method IFont.CloneStyleFrom(IFont) is not working as expected
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
    var f = Font2(); // Based on Font2
    f.IsBold = true;
    return f;
}

var style1 = _workbook.CreateCellStyle();
style1.SetFont(Font1());

var style2 = _workbook.CreateCellStyle();
style2.SetFont(Font2());

var style3 = _workbook.CreateCellStyle();
// Clone style from another, cloning styles work
style3.CloneStyleFrom(style2);
style3.SetFont(Font3());

// COLORS - use NPOI IndexedColors indexes
var ColorBlueIndex = IndexedColors.LightBlue.Index;
var ColorLightGreyIndex = IndexedColors.Grey25Percent.Index;

var style4 = _workbook.CreateCellStyle();
style4.CloneStyleFrom(style2);
// FILLS - must set FillPattern if you want to see something
style4.FillPattern = FillPattern.SolidForeground;
style4.FillForegroundColor = ColorLightGreyIndex;

var style5 = _workbook.CreateCellStyle();
style5.CloneStyleFrom(style4);
style5.FillForegroundColor = ColorBlueIndex;

// Values will be either Even or Odd based on the Age
const string Even = "EVEN";
const string Odd = "ODD";
const string Lastname = "Lastname";
const string Firstname = "Firstname";
const string Age = "Age";

// Set formats to your columns and rows definitions

var mat = Matrix.With()
    .Cols()
        .Col(label: Lastname)
        .Col(label: Firstname, headerCellFormat: new NPOIFormat(style1))
        .Col(label: Age, dataType: DataTypes.Number)
    .Rows()
        .Row(defaultCellFormat: new NPOIFormat(style2))
            .Format(Lastname, new NPOIFormat(style3))
            .Format(Age, new NPOIFormat(style4))
        .Row(key: Odd)
            .Format(Age, new NPOIFormat(style5))
    .RowValues(GetValues()) // GetValues is a method returning List<RowValue> from any business data
    .Build();
```

Same report with ```ClosedXmlFormat```

*Note: Used colors are not exactly the same*

```csharp
Func<IXLStyle, IXLStyle> style1 = (style) =>
{
    style.Font.Italic = true;
    return style;
};

Func<IXLStyle, IXLStyle> style2 = (style) =>
{
    style.Font.FontSize = 12;
    return style;
};

Func<IXLStyle, IXLStyle> style3Partial = (style) =>
{
    style.Font.Bold = true;
    return style;
};
// Use Function Composition in order to inherit styles
// Compose(...) is an Extension method
// Here style3 will be based on style2 and then apply its own style e.g. font is Bold
var style3 = style3Partial.Compose(style2);

var ColorBlueIndex = XLColor.LightBlue;
var ColorLightGreyIndex = XLColor.LightGray;

Func<IXLStyle, IXLStyle> style4Partial = (style) =>
{
    style.Fill.BackgroundColor = ColorLightGreyIndex;
    return style;
};
var style4 = style4Partial.Compose(style2);

Func<IXLStyle, IXLStyle> style5Partial = (style) =>
{
    style.Fill.BackgroundColor = ColorBlueIndex;
    return style;
};
var style5 = style5Partial.Compose(style4);

const string Even = "EVEN";
const string Odd = "ODD";
const string Lastname = "Lastname";
const string Firstname = "Firstname";
const string Age = "Age";

var mat = Matrix.With()
    .Cols()
        .Col(label: Lastname)
        .Col(label: Firstname, headerCellFormat: new ClosedXmlFormat(style1))
        .Col(label: Age, dataType: DataTypes.Number)
    .Rows()
        .Row(defaultCellFormat: new ClosedXmlFormat(style2))
            .Format(Lastname, new ClosedXmlFormat(style3))
            .Format(Age, new ClosedXmlFormat(style4))
        .Row(key: Odd)
            .Format(Age, new ClosedXmlFormat(style5))
    .RowValues(GetValues())
    .Build();

var workbook = new XLWorkbook();
var renderer = new ClosedXmlRenderer(_wb, new ClosedXmlFormatApplier());
using (var fs = File.Create("myReport_using_ClosedXml.xlsx"))
_renderer.GenerateExcelFile(mat, fs);
```
