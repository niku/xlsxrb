// Validates axis fill (solidFill) in spPr on both catAx and valAx.
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

    var catAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().First();
    var catSpPr = catAx.Elements<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault();
    if (catSpPr == null)
        throw new Exception("SCENARIO_FAIL: catAx spPr not found");
    var catFill = catSpPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
    if (catFill == null)
        throw new Exception("SCENARIO_FAIL: catAx solidFill not found");

    var valAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().First();
    var valSpPr = valAx.Elements<DocumentFormat.OpenXml.Drawing.Charts.ChartShapeProperties>().FirstOrDefault();
    if (valSpPr == null)
        throw new Exception("SCENARIO_FAIL: valAx spPr not found");
    var valFill = valSpPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
    if (valFill == null)
        throw new Exception("SCENARIO_FAIL: valAx solidFill not found");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(5).ToList();
    if (errors.Count > 0)
    {
        var errMsg = string.Join("; ", errors.Select(e => e.Description));
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count + " - " + errMsg);
    }

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
