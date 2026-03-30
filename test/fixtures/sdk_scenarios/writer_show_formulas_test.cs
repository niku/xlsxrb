var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var sheetViews = worksheet.SheetViews;

if (sheetViews == null)
    throw new Exception("SCENARIO_FAIL: SheetViews is null");

var sheetView = sheetViews.Elements<SheetView>().First();
if (sheetView.ShowFormulas?.Value != true)
    throw new Exception($"SCENARIO_FAIL: ShowFormulas expected true, got {sheetView.ShowFormulas?.Value}");

Console.WriteLine("SCENARIO_PASS");
