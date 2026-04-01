// Validates that chart category axis has txPr with rot attribute.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var chartPart = workbookPart.WorksheetParts.First().DrawingsPart.ChartParts.First();
    var chart = chartPart.ChartSpace.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();
    var plotArea = chart.PlotArea;

    var catAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().FirstOrDefault()
        ?? throw new Exception("CategoryAxis is missing.");

    var txPr = catAx.Elements<DocumentFormat.OpenXml.Drawing.Charts.TextProperties>().FirstOrDefault()
        ?? throw new Exception("txPr is missing on catAx.");

    var bodyPr = txPr.Elements<DocumentFormat.OpenXml.Drawing.BodyProperties>().FirstOrDefault()
        ?? throw new Exception("BodyProperties is missing in txPr.");

    if (bodyPr.Rotation?.Value != -2700000)
        throw new Exception($"Expected rot=-2700000, got {bodyPr.Rotation?.Value}.");

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
