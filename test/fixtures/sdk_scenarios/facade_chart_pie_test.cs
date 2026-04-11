// Validates that an XLSX generated via Facade API contains a valid pie chart
// and passes OpenXML validation.
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

    var cs = chartParts[0].ChartSpace ?? throw new Exception("ChartSpace is missing.");
    var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
        ?? throw new Exception("Chart element is missing.");

    var plotArea = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().FirstOrDefault()
        ?? throw new Exception("PlotArea is missing.");
    var pieChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChart>().FirstOrDefault()
        ?? throw new Exception("PieChart is missing. Expected pie chart type.");

    var seriesList = pieChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries>().ToList();
    if (seriesList.Count == 0)
        throw new Exception("No series found in pie chart.");

    // Pie chart should NOT have axes
    if (plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().Any())
        throw new Exception("PieChart should not have CategoryAxis.");

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
