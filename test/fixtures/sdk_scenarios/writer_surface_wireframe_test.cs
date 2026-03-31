// Validates wireframe element for surface chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var chartXml = chartPart.ChartSpace.InnerXml;

    if (!chartXml.Contains("wireframe"))
        throw new Exception("SCENARIO_FAIL: wireframe not found in chart XML");
    if (!chartXml.Contains("\"1\""))
        throw new Exception("SCENARIO_FAIL: wireframe val='1' not found");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
