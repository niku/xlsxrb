// Validates axis line_dash (prstDash) in catAx and valAx spPr.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var chartXml = chartPart.ChartSpace.InnerXml;

    if (!chartXml.Contains("prstDash"))
        throw new Exception("SCENARIO_FAIL: prstDash not found in chart XML");

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
