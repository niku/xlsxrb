var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

// Create shared string table with rich text
var ssPart = workbookPart.AddNewPart<SharedStringTablePart>();
var sst = new SharedStringTable { Count = 1, UniqueCount = 1 };

var si = new SharedStringItem();

// Run 0: strike
var run0 = new Run();
var rpr0 = new RunProperties();
rpr0.Append(new Strike());
rpr0.Append(new FontSize { Val = 11 });
rpr0.Append(new RunFont { Val = "Arial" });
run0.Append(rpr0);
run0.Append(new Text("Strike"));
si.Append(run0);

// Run 1: double underline
var run1 = new Run();
var rpr1 = new RunProperties();
rpr1.Append(new Underline { Val = UnderlineValues.Double });
rpr1.Append(new FontSize { Val = 11 });
rpr1.Append(new RunFont { Val = "Arial" });
run1.Append(rpr1);
run1.Append(new Text("DblUnder"));
si.Append(run1);

// Run 2: superscript
var run2 = new Run();
var rpr2 = new RunProperties();
rpr2.Append(new VerticalTextAlignment { Val = VerticalAlignmentRunValues.Superscript });
rpr2.Append(new FontSize { Val = 11 });
rpr2.Append(new RunFont { Val = "Arial" });
run2.Append(rpr2);
run2.Append(new Text("Super"));
si.Append(run2);

// Run 3: theme color with tint, family, scheme
var run3 = new Run();
var rpr3 = new RunProperties();
rpr3.Append(new Color { Theme = 1U, Tint = 0.5 });
rpr3.Append(new FontSize { Val = 11 });
rpr3.Append(new RunFont { Val = "Calibri" });
rpr3.Append(new FontFamily { Val = 2 });
rpr3.Append(new FontScheme { Val = FontSchemeValues.Minor });
run3.Append(rpr3);
run3.Append(new Text("Theme"));
si.Append(run3);

// Run 4: indexed color
var run4 = new Run();
var rpr4 = new RunProperties();
rpr4.Append(new Color { Indexed = 10U });
rpr4.Append(new FontSize { Val = 11 });
rpr4.Append(new RunFont { Val = "Calibri" });
run4.Append(rpr4);
run4.Append(new Text("Indexed"));
si.Append(run4);

sst.Append(si);
ssPart.SharedStringTable = sst;
ssPart.SharedStringTable.Save();

// Create worksheet with cell referencing shared string 0
var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.SharedString,
        CellValue = new CellValue("0") })
);
worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
