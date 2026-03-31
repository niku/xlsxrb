// Validates scatterStyle element on a scatter chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var chartXml = chartPart.ChartSpace.InnerXml;

    if (!chartXml.Contains("scatterStyle"))
        throw new Exception("SCENARIO_FAIL: scatterStyle not found in chart XML");
    if (!chartXml.Contains("\"lineMarker\""))
        throw new Exception("SCENARIO_FAIL: scatterStyle val='lineMarker' not found");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
