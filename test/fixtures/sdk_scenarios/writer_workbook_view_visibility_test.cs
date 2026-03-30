var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var bookViews = workbookPart.Workbook.BookViews;

if (bookViews == null)
    throw new Exception("SCENARIO_FAIL: BookViews is null");

var wbView = bookViews.Elements<WorkbookView>().First();
if (wbView.Visibility?.Value != VisibilityValues.Hidden)
    throw new Exception($"SCENARIO_FAIL: Visibility expected Hidden, got {wbView.Visibility?.Value}");

Console.WriteLine("SCENARIO_PASS");
