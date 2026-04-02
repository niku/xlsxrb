// Validates legendEntry with txPr (font formatting) in legend.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var cs = chartPart.ChartSpace ?? throw new Exception("ChartSpace is missing.");
    var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();

    var legend = chart.Elements<DocumentFormat.OpenXml.Drawing.Charts.Legend>().FirstOrDefault();
    if (legend == null)
        throw new Exception("SCENARIO_FAIL: legend not found");

    var entry = legend.Elements<DocumentFormat.OpenXml.Drawing.Charts.LegendEntry>().FirstOrDefault();
    if (entry == null)
        throw new Exception("SCENARIO_FAIL: legendEntry not found");

    var txPr = entry.Elements<DocumentFormat.OpenXml.Drawing.Charts.TextProperties>().FirstOrDefault();
    if (txPr == null)
        throw new Exception("SCENARIO_FAIL: txPr not found in legendEntry");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(5).ToList();
    if (errors.Count > 0)
    {
        var errMsg = string.Join("; ", errors.Select(e => e.Description));
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count + " - " + errMsg);
    }

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
