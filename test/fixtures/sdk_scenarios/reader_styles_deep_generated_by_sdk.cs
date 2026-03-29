using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var wbPart = doc.AddWorkbookPart();
wbPart.Workbook = new Workbook(new Sheets(
    new Sheet { Id = "rId1", SheetId = 1, Name = "Sheet1" }
));

// Create styles with multiple cellStyleXfs and cellStyles
var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
var stylesheet = new Stylesheet();

var fonts = new Fonts(
    new Font(new FontSize { Val = 11 }, new FontName { Val = "Calibri" }),
    new Font(new Bold(), new FontSize { Val = 14 }, new FontName { Val = "Arial" })
);
fonts.Count = 2;
stylesheet.Append(fonts);

var fills = new Fills(
    new Fill(new PatternFill { PatternType = PatternValues.None }),
    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
);
fills.Count = 2;
stylesheet.Append(fills);

var borders = new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()));
borders.Count = 1;
stylesheet.Append(borders);

var cellStyleXfs = new CellStyleFormats(
    new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0 },
    new CellFormat { NumberFormatId = 0, FontId = 1, FillId = 0, BorderId = 0 }
);
cellStyleXfs.Count = 2;
stylesheet.Append(cellStyleXfs);

var cellXfs = new CellFormats(
    new CellFormat { NumberFormatId = 0, FontId = 0, FillId = 0, BorderId = 0, FormatId = 0 }
);
cellXfs.Count = 1;
stylesheet.Append(cellXfs);

var cellStyles = new CellStyles(
    new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 },
    new CellStyle { Name = "Heading1", FormatId = 1, BuiltinId = 1 }
);
cellStyles.Count = 2;
stylesheet.Append(cellStyles);

stylesPart.Stylesheet = stylesheet;
stylesPart.Stylesheet.Save();

var wsPart = wbPart.AddNewPart<WorksheetPart>("rId1");
var ws = new Worksheet();
var sd = new SheetData();
var row = new Row { RowIndex = 1 };
row.Append(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
    InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("Hello")) });
sd.Append(row);
ws.Append(sd);
wsPart.Worksheet = ws;
wsPart.Worksheet.Save();

doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
