var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var sfp = worksheet.SheetFormatProperties;

if (sfp == null)
    throw new Exception("SCENARIO_FAIL: SheetFormatProperties is null");
if (sfp.DefaultRowHeight?.Value != 18.0)
    throw new Exception($"SCENARIO_FAIL: DefaultRowHeight expected 18.0, got {sfp.DefaultRowHeight?.Value}");
if (sfp.DefaultColumnWidth?.Value != 12.5)
    throw new Exception($"SCENARIO_FAIL: DefaultColWidth expected 12.5, got {sfp.DefaultColumnWidth?.Value}");
if (sfp.BaseColumnWidth?.Value != (uint)10)
    throw new Exception($"SCENARIO_FAIL: BaseColWidth expected 10, got {sfp.BaseColumnWidth?.Value}");
