// Validates legend manual layout element in a chart.
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
        throw new Exception("SCENARIO_FAIL: Legend not found");

    var layout = legend.Elements<DocumentFormat.OpenXml.Drawing.Charts.Layout>().FirstOrDefault();
    if (layout == null)
        throw new Exception("SCENARIO_FAIL: Legend layout not found");

    var ml = layout.Elements<DocumentFormat.OpenXml.Drawing.Charts.ManualLayout>().FirstOrDefault();
    if (ml == null)
        throw new Exception("SCENARIO_FAIL: ManualLayout not found");

    var xEl = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.Left>().FirstOrDefault();
    if (xEl == null)
        throw new Exception("SCENARIO_FAIL: x element not found in ManualLayout");

    var yEl = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.Top>().FirstOrDefault();
    if (yEl == null)
        throw new Exception("SCENARIO_FAIL: y element not found in ManualLayout");

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
