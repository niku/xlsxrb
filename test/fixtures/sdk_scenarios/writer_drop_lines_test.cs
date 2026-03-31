// Validates dropLines and hiLowLines elements on a line chart.
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

    var dropLines = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.DropLines>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: DropLines not found.");

    var hiLowLines = lineChart.Elements<DocumentFormat.OpenXml.Drawing.Charts.HighLowLines>().FirstOrDefault()
        ?? throw new Exception("SCENARIO_FAIL: HighLowLines not found.");

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
