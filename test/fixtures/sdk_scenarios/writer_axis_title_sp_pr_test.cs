// Validates axis title spPr (fill and line) on cat and val axes.
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

    // Find catAx title spPr
    var catAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxis>().First();
    var catTitle = catAx.Title ?? throw new Exception("SCENARIO_FAIL: catAx title is missing");
    var catSpPr = catTitle.ChartShapeProperties;
    if (catSpPr == null)
        throw new Exception("SCENARIO_FAIL: catAx title spPr is missing");
    var catFill = catSpPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
    if (catFill == null)
        throw new Exception("SCENARIO_FAIL: catAx title solidFill is missing");
    var catLn = catSpPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
    if (catLn == null)
        throw new Exception("SCENARIO_FAIL: catAx title ln is missing");

    // Find valAx title spPr
    var valAx = plotArea.Elements<DocumentFormat.OpenXml.Drawing.Charts.ValueAxis>().First();
    var valTitle = valAx.Title ?? throw new Exception("SCENARIO_FAIL: valAx title is missing");
    var valSpPr = valTitle.ChartShapeProperties;
    if (valSpPr == null)
        throw new Exception("SCENARIO_FAIL: valAx title spPr is missing");
    var valFill = valSpPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
    if (valFill == null)
        throw new Exception("SCENARIO_FAIL: valAx title solidFill is missing");
    var valLn = valSpPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
    if (valLn == null)
        throw new Exception("SCENARIO_FAIL: valAx title ln is missing");

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
