// Validates trendline element in a line chart series.
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

    var series = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.LineChartSeries>().First();
    var trendline = series.Elements<DocumentFormat.OpenXml.Drawing.Charts.Trendline>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: Trendline not found.");

    var trendlineType = trendline.Elements<DocumentFormat.OpenXml.Drawing.Charts.TrendlineType>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: TrendlineType not found.");
    if (trendlineType.Val != DocumentFormat.OpenXml.Drawing.Charts.TrendlineValues.Polynomial)
        throw new Exception($"SCENARIO_FAIL: TrendlineType expected poly, got {trendlineType.Val}.");

    var order = trendline.Elements<DocumentFormat.OpenXml.Drawing.Charts.PolynomialOrder>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: Order not found.");
    if (order.Val != 3)
        throw new Exception($"SCENARIO_FAIL: Order expected 3, got {order.Val}.");

    var forward = trendline.Elements<DocumentFormat.OpenXml.Drawing.Charts.Forward>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: Forward not found.");

    var dispRSqr = trendline.Elements<DocumentFormat.OpenXml.Drawing.Charts.DisplayRSquaredValue>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: DisplayRSquaredValue not found.");
    if (dispRSqr.Val != true)
        throw new Exception($"SCENARIO_FAIL: dispRSqr expected true, got {dispRSqr.Val}.");

    var dispEq = trendline.Elements<DocumentFormat.OpenXml.Drawing.Charts.DisplayEquation>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: DisplayEquation not found.");
    if (dispEq.Val != true)
        throw new Exception($"SCENARIO_FAIL: dispEq expected true, got {dispEq.Val}.");

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
