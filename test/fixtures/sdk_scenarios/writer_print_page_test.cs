var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;

// Check printOptions
var po = worksheet.GetFirstChild<PrintOptions>();
if (po == null)
    throw new Exception("SCENARIO_FAIL: PrintOptions is null");
if (po.GridLines?.Value != true)
    throw new Exception($"SCENARIO_FAIL: GridLines expected true, got {po.GridLines?.Value}");

// Check pageMargins
var pm = worksheet.GetFirstChild<PageMargins>();
if (pm == null)
    throw new Exception("SCENARIO_FAIL: PageMargins is null");
if (Math.Abs(pm.Left.Value - 0.7) > 0.01)
    throw new Exception($"SCENARIO_FAIL: Left margin expected 0.7, got {pm.Left.Value}");

// Check pageSetup
var ps = worksheet.GetFirstChild<PageSetup>();
if (ps == null)
    throw new Exception("SCENARIO_FAIL: PageSetup is null");
if (ps.Orientation?.Value != OrientationValues.Landscape)
    throw new Exception($"SCENARIO_FAIL: Orientation expected Landscape, got {ps.Orientation?.Value}");

// Check headerFooter
var hf = worksheet.GetFirstChild<HeaderFooter>();
if (hf == null)
    throw new Exception("SCENARIO_FAIL: HeaderFooter is null");
if (hf.OddHeader?.Text != "&CPage &P")
    throw new Exception($"SCENARIO_FAIL: OddHeader expected '&CPage &P', got '{hf.OddHeader?.Text}'");

// Check rowBreaks
var rb = worksheet.GetFirstChild<RowBreaks>();
if (rb == null)
    throw new Exception("SCENARIO_FAIL: RowBreaks is null");
var brkIds = rb.Elements<Break>().Select(b => (int)b.Id.Value).ToList();
if (!brkIds.SequenceEqual(new[] { 10, 20 }))
    throw new Exception($"SCENARIO_FAIL: RowBreaks expected [10,20], got [{string.Join(",", brkIds)}]");
