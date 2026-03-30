var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var sheetProps = worksheet.SheetProperties;

if (sheetProps == null)
    throw new Exception("SCENARIO_FAIL: SheetProperties is null");

if (sheetProps.CodeName?.Value != "MySheet")
    throw new Exception($"SCENARIO_FAIL: CodeName expected MySheet, got {sheetProps.CodeName?.Value}");

Console.WriteLine("SCENARIO_PASS");
