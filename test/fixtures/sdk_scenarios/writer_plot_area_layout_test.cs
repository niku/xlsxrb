// Validates plot area manual layout in plotArea > layout > manualLayout.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var cs = chartPart.ChartSpace ?? throw new Exception("ChartSpace is missing.");
    var chart = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();
    var plotArea = chart.PlotArea;

    var layout = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.Layout>().FirstOrDefault();
    if (layout == null)
        throw new Exception("SCENARIO_FAIL: plotArea layout not found");

    var ml = layout.Elements<DocumentFormat.OpenXml.Drawing.Charts.ManualLayout>().FirstOrDefault();
    if (ml == null)
        throw new Exception("SCENARIO_FAIL: manualLayout not found");

    var lt = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.LayoutTarget>().FirstOrDefault();
    if (lt == null || lt.Val == null)
        throw new Exception("SCENARIO_FAIL: layoutTarget not found");

    var x = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.Left>().FirstOrDefault();
    if (x == null)
        throw new Exception("SCENARIO_FAIL: x element not found");

    var y = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.Top>().FirstOrDefault();
    if (y == null)
        throw new Exception("SCENARIO_FAIL: y element not found");

    var w = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.Width>().FirstOrDefault();
    if (w == null)
        throw new Exception("SCENARIO_FAIL: w element not found");

    var h = ml.Elements<DocumentFormat.OpenXml.Drawing.Charts.Height>().FirstOrDefault();
    if (h == null)
        throw new Exception("SCENARIO_FAIL: h element not found");

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
