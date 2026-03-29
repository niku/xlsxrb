// Validates pivot table retention.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
        ?? throw new Exception("No sheets found.");

    if (sheets.Count < 2)
        throw new Exception("Expected at least 2 sheets after retention.");

    var pivotSheet = (WorksheetPart)workbookPart.GetPartById(sheets[1].Id.Value);
    var pivotParts = pivotSheet.PivotTableParts.ToList();
    if (pivotParts.Count == 0)
        throw new Exception("No pivot table parts found after retention.");
}
finally
{
    document.Dispose();
}
