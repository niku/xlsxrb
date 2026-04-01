// Validates that an XLSX chart series has a:ln with cap attribute and a:round join.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    var chartPart = drawingsPart.ChartParts.FirstOrDefault()
        ?? throw new Exception("ChartPart is missing.");

    var chart = chartPart.ChartSpace?.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
        ?? throw new Exception("Chart is missing.");

    var plotArea = chart.PlotArea ?? throw new Exception("PlotArea is missing.");
    var lineChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().FirstOrDefault()
        ?? throw new Exception("LineChart is missing.");

    var ser = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>().FirstOrDefault()
        ?? throw new Exception("Series is missing.");

    var spPr = ser.Elements<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault()
        ?? throw new Exception("ChartShapeProperties is missing.");

    var ln = spPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault()
        ?? throw new Exception("Outline is missing.");

    if (ln.Width?.Value != 25400) throw new Exception($"Expected w=25400, got {ln.Width?.Value}.");
    if (ln.CapType?.Value != DocumentFormat.OpenXml.Drawing.LineCapValues.Round)
        throw new Exception($"Expected cap=rnd, got {ln.CapType?.Value}.");

    var round = ln.Elements<DocumentFormat.OpenXml.Drawing.Round>().FirstOrDefault()
        ?? throw new Exception("Round join is missing.");

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
