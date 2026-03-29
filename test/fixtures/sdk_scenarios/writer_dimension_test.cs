var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var dimension = worksheet.SheetDimension;

if (dimension == null)
    throw new Exception("SCENARIO_FAIL: SheetDimension is null");
if (dimension.Reference?.Value != "B2:D5")
    throw new Exception($"SCENARIO_FAIL: Dimension expected B2:D5, got {dimension.Reference?.Value}");
