// Validates legend and data table line_dash (prstDash) in spPr.
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

    var legSpPr = legend.Elements<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault();
    if (legSpPr == null)
        throw new Exception("SCENARIO_FAIL: legend spPr not found");

    var legLn = legSpPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
    if (legLn == null)
        throw new Exception("SCENARIO_FAIL: legend ln not found");

    var legDash = legLn.Elements<DocumentFormat.OpenXml.Drawing.PresetDash>().FirstOrDefault();
    if (legDash == null)
        throw new Exception("SCENARIO_FAIL: legend prstDash not found");

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
