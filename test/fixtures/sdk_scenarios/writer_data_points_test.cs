// Validates dPt elements for series data points on a pie chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();

    var pieChart = chartPart.ChartSpace
        .Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First()
        .Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().First()
        .Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChart>().First();

    var series = pieChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries>().First();
    var dpts = series.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataPoint>().ToList();
    if (dpts.Count != 3)
        throw new Exception($"SCENARIO_FAIL: Expected 3 dPt elements, found {dpts.Count}.");

    var expectedColors = new[] { "FF0000", "00FF00", "0000FF" };
    for (int i = 0; i < 3; i++)
    {
        var dp = dpts[i];
        var idx = dp.Elements<DocumentFormat.OpenXml.Drawing.Charts.Index>().FirstOrDefault()
            ?? throw new Exception($"SCENARIO_FAIL: dPt[{i}] missing Index.");
        if (idx.Val != (uint)i)
            throw new Exception($"SCENARIO_FAIL: dPt[{i}] Index expected {i}, got {idx.Val}.");

        var spPr = dp.Elements<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault()
            ?? throw new Exception($"SCENARIO_FAIL: dPt[{i}] missing spPr.");
        var solidFill = spPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault()
            ?? throw new Exception($"SCENARIO_FAIL: dPt[{i}] missing solidFill.");
        var srgbClr = solidFill.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()
            ?? throw new Exception($"SCENARIO_FAIL: dPt[{i}] missing srgbClr.");
        if (srgbClr.Val != expectedColors[i])
            throw new Exception($"SCENARIO_FAIL: dPt[{i}] color expected {expectedColors[i]}, got {srgbClr.Val}.");
    }

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
