// Validates showLeaderLines and leaderLines with spPr in dLbls for a pie chart.
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

    var pieChart = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChart>().FirstOrDefault();
    if (pieChart == null)
        throw new Exception("SCENARIO_FAIL: PieChart not found");

    var ser = pieChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.PieChartSeries>().First();
    var dLbls = ser.Elements<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault()
        ?? plotArea.Descendants<DocumentFormat.OpenXml.Drawing.Charts.DataLabels>().FirstOrDefault();
    if (dLbls == null)
        throw new Exception("SCENARIO_FAIL: DataLabels not found");

    var showLL = dLbls.Elements<DocumentFormat.OpenXml.Drawing.Charts.ShowLeaderLines>().FirstOrDefault();
    if (showLL == null || showLL.Val == null || !showLL.Val.Value)
        throw new Exception("SCENARIO_FAIL: showLeaderLines not found or not true");

    var ll = dLbls.Elements<DocumentFormat.OpenXml.Drawing.Charts.LeaderLines>().FirstOrDefault();
    if (ll == null)
        throw new Exception("SCENARIO_FAIL: leaderLines not found");

    var llSpPr = ll.ChartShapeProperties;
    if (llSpPr == null)
        throw new Exception("SCENARIO_FAIL: leaderLines spPr not found");

    var ln = llSpPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
    if (ln == null)
        throw new Exception("SCENARIO_FAIL: leaderLines ln not found");

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
