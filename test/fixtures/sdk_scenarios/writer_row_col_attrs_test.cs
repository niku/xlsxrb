var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

// Check row outline level and collapsed
var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>().ToList();
var row2 = rows.FirstOrDefault(r => r.RowIndex?.Value == 2);
if (row2 == null)
    throw new Exception("SCENARIO_FAIL: Row 2 not found");
if (row2.OutlineLevel?.Value != 1)
    throw new Exception($"SCENARIO_FAIL: Row2 OutlineLevel expected 1, got {row2.OutlineLevel?.Value}");

var row3 = rows.FirstOrDefault(r => r.RowIndex?.Value == 3);
if (row3 == null)
    throw new Exception("SCENARIO_FAIL: Row 3 not found");
if (row3.Collapsed?.Value != true)
    throw new Exception($"SCENARIO_FAIL: Row3 Collapsed expected true, got {row3.Collapsed?.Value}");

// Check column attributes
var cols = worksheet.GetFirstChild<Columns>().Elements<Column>().ToList();
var colB = cols.FirstOrDefault(c => c.Min?.Value == 2 && c.Max?.Value == 2);
if (colB == null)
    throw new Exception("SCENARIO_FAIL: Column B not found");
if (colB.Hidden?.Value != true)
    throw new Exception($"SCENARIO_FAIL: ColB Hidden expected true, got {colB.Hidden?.Value}");

var colC = cols.FirstOrDefault(c => c.Min?.Value == 3 && c.Max?.Value == 3);
if (colC == null)
    throw new Exception("SCENARIO_FAIL: Column C not found");
if (colC.OutlineLevel?.Value != 2)
    throw new Exception($"SCENARIO_FAIL: ColC OutlineLevel expected 2, got {colC.OutlineLevel?.Value}");
if (colC.Collapsed?.Value != true)
    throw new Exception($"SCENARIO_FAIL: ColC Collapsed expected true, got {colC.Collapsed?.Value}");
