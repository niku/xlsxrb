var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);
    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    var chartParts = drawingsPart.ChartParts.ToList();
    if (chartParts.Count < 4) throw new Exception($"Expected at least 4 chart parts but got {chartParts.Count}.");

    // Chart 1: area
    var cs1 = chartParts[0].ChartSpace ?? throw new Exception("ChartSpace missing.");
    var chart1 = cs1.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();
    var plotArea1 = chart1.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().First();
    var areaChart = plotArea1.Elements<DocumentFormat.OpenXml.Drawing.Charts.AreaChart>().FirstOrDefault();
    if (areaChart == null) throw new Exception("AreaChart not found in chart 1.");

    // Chart 2: scatter
    var cs2 = chartParts[1].ChartSpace ?? throw new Exception("ChartSpace missing.");
    var chart2 = cs2.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();
    var plotArea2 = chart2.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().First();
    var scatterChart = plotArea2.Elements<DocumentFormat.OpenXml.Drawing.Charts.ScatterChart>().FirstOrDefault();
    if (scatterChart == null) throw new Exception("ScatterChart not found in chart 2.");

    // Chart 3: doughnut
    var cs3 = chartParts[2].ChartSpace ?? throw new Exception("ChartSpace missing.");
    var chart3 = cs3.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();
    var plotArea3 = chart3.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().First();
    var doughnutChart = plotArea3.Elements<DocumentFormat.OpenXml.Drawing.Charts.DoughnutChart>().FirstOrDefault();
    if (doughnutChart == null) throw new Exception("DoughnutChart not found in chart 3.");
    // Doughnut should not have axes
    if (plotArea3.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().Any())
        throw new Exception("DoughnutChart should not have CategoryAxis.");

    // Chart 4: radar
    var cs4 = chartParts[3].ChartSpace ?? throw new Exception("ChartSpace missing.");
    var chart4 = cs4.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().First();
    var plotArea4 = chart4.Elements<DocumentFormat.OpenXml.Drawing.Charts.PlotArea>().First();
    var radarChart = plotArea4.Elements<DocumentFormat.OpenXml.Drawing.Charts.RadarChart>().FirstOrDefault();
    if (radarChart == null) throw new Exception("RadarChart not found in chart 4.");
}
finally
{
    document.Dispose();
}
