// Validates that an XLSX chart contains series with line formatting (a:ln inside c:spPr).
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

    var plotArea = chart.PlotArea ?? throw new Exception("PlotArea is missing.");
    var lineChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().FirstOrDefault()
        ?? throw new Exception("LineChart is missing.");

    var series = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>().ToList();
    if (series.Count == 0)
        throw new Exception("No series found.");

    var ser = series[0];
    var spPr = ser.Elements<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault()
        ?? throw new Exception("Series ShapeProperties is missing.");

    var ln = spPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault()
        ?? throw new Exception("Outline (a:ln) is missing in series spPr.");

    if (ln.Width == null || ln.Width.Value != 25400)
        throw new Exception($"Expected line width 25400 but got {ln.Width?.Value}.");

    var solidFill = ln.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault()
        ?? throw new Exception("SolidFill inside ln is missing.");

    var srgbClr = solidFill.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()
        ?? throw new Exception("RgbColorModelHex inside solidFill is missing.");

    if (srgbClr.Val?.Value != "0000FF")
        throw new Exception($"Expected line color 0000FF but got {srgbClr.Val?.Value}.");

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
