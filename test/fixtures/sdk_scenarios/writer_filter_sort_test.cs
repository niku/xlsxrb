var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

var af = worksheet.GetFirstChild<AutoFilter>();
if (af == null)
    throw new Exception("SCENARIO_FAIL: AutoFilter is null");
if (af.Reference?.Value != "A1:C10")
    throw new Exception($"SCENARIO_FAIL: AutoFilter ref expected A1:C10, got {af.Reference?.Value}");

var fcs = af.Elements<FilterColumn>().ToList();
if (fcs.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 FilterColumns, got {fcs.Count}");

// First: filters with values
var fc0 = fcs.First(f => f.ColumnId?.Value == 0);
var filters = fc0.GetFirstChild<Filters>();
if (filters == null)
    throw new Exception("SCENARIO_FAIL: Filters element is null for col 0");
var vals = filters.Elements<Filter>().Select(f => f.Val?.Value).ToList();
if (!vals.SequenceEqual(new[] { "A", "B" }))
    throw new Exception($"SCENARIO_FAIL: Filter values expected [A,B], got [{string.Join(",", vals)}]");

// Second: custom filter
var fc1 = fcs.First(f => f.ColumnId?.Value == 1);
var customs = fc1.GetFirstChild<CustomFilters>();
if (customs == null)
    throw new Exception("SCENARIO_FAIL: CustomFilters is null for col 1");
var cf = customs.Elements<CustomFilter>().First();
if (cf.Operator?.Value != FilterOperatorValues.GreaterThan)
    throw new Exception($"SCENARIO_FAIL: CustomFilter operator expected GreaterThan, got {cf.Operator?.Value}");
if (cf.Val?.Value != "100")
    throw new Exception($"SCENARIO_FAIL: CustomFilter val expected 100, got {cf.Val?.Value}");

// Check sortState
var ss = worksheet.GetFirstChild<SortState>();
if (ss == null)
    throw new Exception("SCENARIO_FAIL: SortState is null");
if (ss.Reference?.Value != "A1:B10")
    throw new Exception($"SCENARIO_FAIL: SortState ref expected A1:B10, got {ss.Reference?.Value}");
var sortConds = ss.Elements<SortCondition>().ToList();
if (sortConds.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 SortConditions, got {sortConds.Count}");
if (sortConds[1].Descending?.Value != true)
    throw new Exception($"SCENARIO_FAIL: SortCondition[1] descending expected true, got {sortConds[1].Descending?.Value}");
