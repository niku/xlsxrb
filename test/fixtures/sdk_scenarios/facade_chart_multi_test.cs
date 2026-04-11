// Validates that an XLSX generated via Facade API contains multiple charts
// on the same sheet and passes OpenXML validation.
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
    if (chartParts.Count < 2)
        throw new Exception($"Expected at least 2 chart parts, got {chartParts.Count}.");

    // Verify each chart has valid structure
    foreach (var cp in chartParts)
    {
        var cs = cp.ChartSpace ?? throw new Exception("ChartSpace is missing in one of the charts.");
        var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
            ?? throw new Exception("Chart element is missing in one of the charts.");
        var plotArea = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault()
            ?? throw new Exception("PlotArea is missing in one of the charts.");
    }

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri}, Path: {e.Path?.XPath})"));
        throw new Exception($"OpenXML Validation errors:\n{messages}");
    }
}
finally
{
    document.Dispose();
}
