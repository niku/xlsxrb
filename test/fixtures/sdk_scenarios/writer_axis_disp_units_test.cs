// Validates dispUnits and builtInUnit elements on chart value axis.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var chartXml = chartPart.ChartSpace.InnerXml;

    if (!chartXml.Contains("dispUnits"))
        throw new Exception("SCENARIO_FAIL: dispUnits not found in chart XML");
    if (!chartXml.Contains("builtInUnit"))
        throw new Exception("SCENARIO_FAIL: builtInUnit not found in chart XML");
    if (!chartXml.Contains("\"thousands\""))
        throw new Exception("SCENARIO_FAIL: builtInUnit val='thousands' not found");

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
