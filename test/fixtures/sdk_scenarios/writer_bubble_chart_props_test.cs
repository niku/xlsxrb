// Validates bubbleScale, showNegBubbles, and sizeRepresents elements in bubble chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var chartXml = chartPart.ChartSpace.InnerXml;

    if (!chartXml.Contains("bubbleScale"))
        throw new Exception("SCENARIO_FAIL: bubbleScale not found in chart XML");
    if (!chartXml.Contains("bubble3D"))
        throw new Exception("SCENARIO_FAIL: bubble3D not found in chart XML");
    if (!chartXml.Contains("showNegBubbles"))
        throw new Exception("SCENARIO_FAIL: showNegBubbles not found in chart XML");
    if (!chartXml.Contains("sizeRepresents"))
        throw new Exception("SCENARIO_FAIL: sizeRepresents not found in chart XML");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
