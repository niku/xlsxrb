var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

// Styles with gradient fill using theme/indexed stop colors
var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
var stylesheet = new Stylesheet();

var fonts = new Fonts(new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" }));
fonts.Count = 1;
stylesheet.Append(fonts);

var fills = new Fills();
fills.Append(new Fill(new PatternFill { PatternType = PatternValues.None }));
fills.Append(new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
// Gradient fill with theme and indexed stops
var gradFill = new GradientFill { Degree = 90 };
gradFill.Append(new GradientStop { Position = 0, Color = new Color { Theme = 4U, Tint = -0.5 } });
gradFill.Append(new GradientStop { Position = 1, Color = new Color { Indexed = 12U } });
fills.Append(new Fill(gradFill));
fills.Count = 3;
stylesheet.Append(fills);

var borders = new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()));
borders.Count = 1;
stylesheet.Append(borders);

var cellStyleFormats = new CellStyleFormats(new CellFormat { FontId = 0, FillId = 0, BorderId = 0 });
cellStyleFormats.Count = 1;
stylesheet.Append(cellStyleFormats);

var cellFormats = new CellFormats();
cellFormats.Append(new CellFormat { FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 });
cellFormats.Append(new CellFormat { FontId = 0, FillId = 2, BorderId = 0, FormatId = 0, ApplyFill = true });
cellFormats.Count = 2;
stylesheet.Append(cellFormats);

var cellStyles = new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 });
cellStyles.Count = 1;
stylesheet.Append(cellStyles);

stylesPart.Stylesheet = stylesheet;
stylesPart.Stylesheet.Save();

// Worksheet
var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", StyleIndex = 1U, DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("themed gradient")) })
);
worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
