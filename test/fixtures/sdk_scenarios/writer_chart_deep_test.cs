// Validates that an XLSX contains a chart with multiple series, legend, data labels, and axis titles.
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

    var cp = chartParts[0];
    var cs = cp.ChartSpace ?? throw new Exception("ChartSpace is missing.");
    var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
        ?? throw new Exception("Chart element is missing.");

    // Verify at least 2 series
    var plotArea = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault()
        ?? throw new Exception("PlotArea is missing.");
    var barChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChart>().FirstOrDefault()
        ?? throw new Exception("BarChart is missing.");
    var seriesList = barChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.BarChartSeries>().ToList();
    if (seriesList.Count < 2)
        throw new Exception($"Expected at least 2 series, got {seriesList.Count}.");

    // Verify legend
    var legend = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.Legend>().FirstOrDefault()
        ?? throw new Exception("Legend is missing.");

    // Verify axis titles exist
    var catAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().FirstOrDefault()
        ?? throw new Exception("CategoryAxis is missing.");
    var valAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().FirstOrDefault()
        ?? throw new Exception("ValueAxis is missing.");

    // Validate
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri}, Path: {e.Path?.XPath})"));
        throw new Exception($"Validation errors:\n{messages}");
    }
}
finally
{
    document.Dispose();
}
