// Validates that an XLSX contains a chart in the first sheet's drawing.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    var chartParts = drawingsPart.ChartParts.ToList();
    if (chartParts.Count == 0)
        throw new Exception("No chart parts found in drawing.");

    // Verify chart has a chart space.
    foreach (var cp in chartParts)
    {
        var cs = cp.ChartSpace ?? throw new Exception("ChartSpace is missing.");
        var c = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
            ?? throw new Exception("Chart element is missing.");
    }
}
finally
{
    document.Dispose();
}
