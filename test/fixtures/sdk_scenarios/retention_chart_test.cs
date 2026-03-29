// Validates chart retention: XLSX should still contain a chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing after retention.");

    var chartParts = drawingsPart.ChartParts.ToList();
    if (chartParts.Count == 0)
        throw new Exception("No chart parts found after retention.");
}
finally
{
    document.Dispose();
}
