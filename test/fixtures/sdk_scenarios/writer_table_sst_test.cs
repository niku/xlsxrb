var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

// Check tableParts in worksheet
var tableParts = worksheetPart.Worksheet.GetFirstChild<TableParts>();
if (tableParts == null)
    throw new Exception("SCENARIO_FAIL: TableParts is null");
if (tableParts.Count?.Value != 1)
    throw new Exception($"SCENARIO_FAIL: Expected 1 tablePart, got {tableParts.Count?.Value}");

// Check table part
var tables = worksheetPart.TableDefinitionParts.ToList();
if (tables.Count != 1)
    throw new Exception($"SCENARIO_FAIL: Expected 1 TableDefinitionPart, got {tables.Count}");
var table = tables[0].Table;
if (table.Reference?.Value != "A1:B5")
    throw new Exception($"SCENARIO_FAIL: Table ref expected A1:B5, got {table.Reference?.Value}");
var cols = table.TableColumns.Elements<TableColumn>().ToList();
if (cols.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 columns, got {cols.Count}");
if (cols[0].Name?.Value != "Name")
    throw new Exception($"SCENARIO_FAIL: Col 0 name expected Name, got {cols[0].Name?.Value}");
if (cols[1].Name?.Value != "Age")
    throw new Exception($"SCENARIO_FAIL: Col 1 name expected Age, got {cols[1].Name?.Value}");

// Check shared strings
var sstPart = workbookPart.SharedStringTablePart;
if (sstPart == null)
    throw new Exception("SCENARIO_FAIL: SharedStringTablePart is null");
var sst = sstPart.SharedStringTable;
var items = sst.Elements<SharedStringItem>().ToList();
if (items.Count < 2)
    throw new Exception($"SCENARIO_FAIL: Expected at least 2 SST items, got {items.Count}");

// Check cells use shared string references
var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
var cellA1 = sheetData.Elements<Row>().First().Elements<Cell>().First();
if (cellA1.DataType?.Value != CellValues.SharedString)
    throw new Exception($"SCENARIO_FAIL: A1 DataType expected SharedString, got {cellA1.DataType?.Value}");

doc.Dispose();
