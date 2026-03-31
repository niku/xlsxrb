// Validates that an XLSX line chart contains series with marker (symbol + size).
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
    if (chartParts.Count == 0) throw new Exception("No chart parts found.");

    var cp = chartParts[0];
    var cs = cp.ChartSpace ?? throw new Exception("ChartSpace is missing.");
    var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
        ?? throw new Exception("Chart element is missing.");
    var plotArea = chart.PlotArea ?? throw new Exception("PlotArea is missing.");
    var lineChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().FirstOrDefault()
        ?? throw new Exception("LineChart is missing.");

    var series = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>().ToList();
    if (series.Count == 0) throw new Exception("No series found.");

    var marker = series[0].Elements<DocumentFormat.OpenXml.Drawing.Charts.Marker>().FirstOrDefault()
        ?? throw new Exception("Series Marker is missing.");

    var symbol = marker.Elements<DocumentFormat.OpenXml.Drawing.Charts.Symbol>().FirstOrDefault()
        ?? throw new Exception("Marker Symbol is missing.");
    if (symbol.Val?.InnerText != "diamond")
        throw new Exception($"Expected symbol 'diamond' but got '{symbol.Val?.InnerText}'.");

    var size = marker.Elements<DocumentFormat.OpenXml.Drawing.Charts.Size>().FirstOrDefault()
        ?? throw new Exception("Marker Size is missing.");
    if (size.Val?.Value != 8)
        throw new Exception($"Expected size 8 but got {size.Val?.Value}.");

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
