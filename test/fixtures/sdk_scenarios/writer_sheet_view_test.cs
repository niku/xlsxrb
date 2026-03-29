var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>().First();
var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
var worksheet = worksheetPart.Worksheet;
var sheetViews = worksheet.SheetViews;

if (sheetViews == null)
    throw new Exception("SCENARIO_FAIL: SheetViews is null");

var sheetView = sheetViews.Elements<SheetView>().First();
if (sheetView.ShowGridLines?.Value != false)
    throw new Exception($"SCENARIO_FAIL: ShowGridLines expected false, got {sheetView.ShowGridLines?.Value}");
if (sheetView.ZoomScale?.Value != (uint)150)
    throw new Exception($"SCENARIO_FAIL: ZoomScale expected 150, got {sheetView.ZoomScale?.Value}");

var pane = sheetView.Pane;
if (pane == null)
    throw new Exception("SCENARIO_FAIL: Pane is null");
if (pane.VerticalSplit?.Value != 1.0)
    throw new Exception($"SCENARIO_FAIL: VerticalSplit expected 1, got {pane.VerticalSplit?.Value}");
if (pane.HorizontalSplit?.Value != 1.0)
    throw new Exception($"SCENARIO_FAIL: HorizontalSplit expected 1, got {pane.HorizontalSplit?.Value}");
if (pane.State?.Value != PaneStateValues.Frozen)
    throw new Exception($"SCENARIO_FAIL: PaneState expected Frozen, got {pane.State?.Value}");

var selection = sheetView.Elements<Selection>().FirstOrDefault();
if (selection == null)
    throw new Exception("SCENARIO_FAIL: Selection is null");
if (selection.ActiveCell?.Value != "C5")
    throw new Exception($"SCENARIO_FAIL: ActiveCell expected C5, got {selection.ActiveCell?.Value}");
if (selection.SequenceOfReferences?.InnerText != "C5")
    throw new Exception($"SCENARIO_FAIL: Sqref expected C5, got {selection.SequenceOfReferences?.InnerText}");
