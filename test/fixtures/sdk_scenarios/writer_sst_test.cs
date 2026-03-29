var doc = SpreadsheetDocument.Open(XlsxPath, false);
var workbookPart = doc.WorkbookPart;
var sstPart = workbookPart.SharedStringTablePart;
if (sstPart == null)
    throw new Exception("SCENARIO_FAIL: SharedStringTablePart is null");
var items = sstPart.SharedStringTable.Elements<SharedStringItem>().ToList();
if (items.Count < 1)
    throw new Exception($"SCENARIO_FAIL: Expected at least 1 SST entry, got {items.Count}");
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var firstCell = wsPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().First().Elements<Cell>().First();
if (firstCell.DataType?.Value != CellValues.SharedString)
    throw new Exception($"SCENARIO_FAIL: Expected SharedString type, got {firstCell.DataType?.Value}");
doc.Dispose();
