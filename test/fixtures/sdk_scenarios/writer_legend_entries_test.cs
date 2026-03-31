// Validates that an XLSX chart contains legend entries with delete attribute.
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

    // Verify legend with entries
    var legend = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.Legend>().FirstOrDefault()
        ?? throw new Exception("Legend is missing.");
    var legendEntries = legend.Elements<DocumentFormat.OpenXml.Drawing.Charts.LegendEntry>().ToList();
    if (legendEntries.Count == 0)
        throw new Exception("No legend entries found.");

    var firstEntry = legendEntries[0];
    var idx = firstEntry.Elements<DocumentFormat.OpenXml.Drawing.Charts.Index>().FirstOrDefault()
        ?? throw new Exception("LegendEntry Index is missing.");
    var del = firstEntry.Elements<DocumentFormat.OpenXml.Drawing.Charts.Delete>().FirstOrDefault()
        ?? throw new Exception("LegendEntry Delete is missing.");

    if (del.Val == null || !del.Val.Value)
        throw new Exception("LegendEntry Delete should be true.");

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
