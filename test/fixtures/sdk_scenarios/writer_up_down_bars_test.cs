// Validates upDownBars with gapWidth on a line chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();

    var lineChart = chartPart.ChartSpace
        .Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First()
        .Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().First()
        .Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChart>().First();

    var upDownBars = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.UpDownBars>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: UpDownBars not found.");

    var gw = upDownBars.Elements<DocumentFormat.OpenXml.Drawing.Charts.GapWidth>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: GapWidth not found in UpDownBars.");
    if (gw.Val != 150) throw new Exception($"SCENARIO_FAIL: Expected GapWidth=150, got {gw.Val}.");

    var upBars = upDownBars.Elements<DocumentFormat.OpenXml.Drawing.Charts.UpBars>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: UpBars not found.");
    var downBars = upDownBars.Elements<DocumentFormat.OpenXml.Drawing.Charts.DownBars>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: DownBars not found.");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(5).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri}, Path: {e.Path?.XPath})"));
        throw new Exception($"SCENARIO_FAIL: validation errors:\n{messages}");
    }

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
