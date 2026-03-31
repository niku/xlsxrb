// Validates that view3D element with rotX/rotY is correctly written for a 3D bar chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var chartXml = chartPart.ChartSpace.InnerXml;

    if (!chartXml.Contains("view3D"))
        throw new Exception("SCENARIO_FAIL: view3D element not found in chart XML");
    if (!chartXml.Contains("rotX"))
        throw new Exception("SCENARIO_FAIL: rotX element not found in chart XML");
    if (!chartXml.Contains("rotY"))
        throw new Exception("SCENARIO_FAIL: rotY element not found in chart XML");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(5).ToList();
    if (errors.Count > 0)
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
