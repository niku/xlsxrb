var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var sheetProps = worksheet.SheetProperties;

if (sheetProps == null)
    throw new Exception("SCENARIO_FAIL: SheetProperties is null");

var tabColor = sheetProps.TabColor;
if (tabColor == null)
    throw new Exception("SCENARIO_FAIL: TabColor is null");
if (tabColor.Rgb?.Value != "FF0000FF")
    throw new Exception($"SCENARIO_FAIL: TabColor expected FF0000FF, got {tabColor.Rgb?.Value}");

var outlinePr = sheetProps.OutlineProperties;
if (outlinePr == null)
    throw new Exception("SCENARIO_FAIL: OutlineProperties is null");
if (outlinePr.SummaryBelow?.Value != false)
    throw new Exception($"SCENARIO_FAIL: SummaryBelow expected false, got {outlinePr.SummaryBelow?.Value}");
if (outlinePr.SummaryRight?.Value != true)
    throw new Exception($"SCENARIO_FAIL: SummaryRight expected true, got {outlinePr.SummaryRight?.Value}");

Console.WriteLine("SCENARIO_PASS");
