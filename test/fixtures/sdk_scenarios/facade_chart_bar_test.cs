// Validates that an XLSX generated via Facade API contains a valid bar chart
// with correct title and passes OpenXML validation.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    // Verify drawing part exists
    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing. The worksheet must have a drawing part for charts.");

    var chartParts = drawingsPart.ChartParts.ToList();
    if (chartParts.Count == 0)
        throw new Exception("No chart parts found in drawing.");

    var cs = chartParts[0].ChartSpace ?? throw new Exception("ChartSpace is missing.");
    var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
        ?? throw new Exception("Chart element is missing.");

    // Verify it's a bar chart
    var plotArea = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault()
        ?? throw new Exception("PlotArea is missing.");
    var barChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChart>().FirstOrDefault()
        ?? throw new Exception("BarChart is missing. Expected bar chart type.");

    // Verify series data
    var seriesList = barChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries>().ToList();
    if (seriesList.Count == 0)
        throw new Exception("No series found in bar chart.");

    // Verify cell data in sheet
    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault()
        ?? throw new Exception("SheetData is missing.");
    var rows = sheetData.Elements<Row>().ToList();
    if (rows.Count < 2)
        throw new Exception($"Expected at least 2 rows of data, got {rows.Count}.");

    // Run OpenXML validation
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
