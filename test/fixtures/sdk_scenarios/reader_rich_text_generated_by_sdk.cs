// Creates an XLSX with rich text in shared strings.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var ssPart = workbookPart.AddNewPart<SharedStringTablePart>();
var sst = new SharedStringTable();
var si = new SharedStringItem();
// Run 1: bold "Hello"
var run1 = new Run();
var rpr1 = new RunProperties();
rpr1.Append(new Bold());
rpr1.Append(new FontSize { Val = 14 });
run1.Append(rpr1);
run1.Append(new DocumentFormat.OpenXml.Spreadsheet.Text("Hello"));
si.Append(run1);
// Run 2: normal " World"
var run2 = new Run();
run2.Append(new DocumentFormat.OpenXml.Spreadsheet.Text(" World"));
si.Append(run2);
sst.Append(si);
sst.Count = 1;
sst.UniqueCount = 1;
ssPart.SharedStringTable = sst;
ssPart.SharedStringTable.Save();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.SharedString, CellValue = new CellValue("0") })
);
worksheetPart.Worksheet = new Worksheet(sheetData);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
