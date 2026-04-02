// Validates chart space spPr (background fill and border line) on a chart.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var chartPart = wsPart.DrawingsPart.ChartParts.First();
    var cs = chartPart.ChartSpace ?? throw new Exception("ChartSpace is missing.");

    // Check that spPr exists directly under chartSpace
    var spPr = cs.Elements<DocumentFormat.OpenXml.Drawing.Charts.ShapeProperties>().FirstOrDefault();
    if (spPr == null)
        throw new Exception("SCENARIO_FAIL: ShapeProperties (spPr) not found in chartSpace");

    // Check solidFill exists
    var solidFill = spPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault();
    if (solidFill == null)
        throw new Exception("SCENARIO_FAIL: solidFill not found in chartSpace spPr");

    // Check ln exists
    var ln = spPr.Elements<DocumentFormat.OpenXml.Drawing.Outline>().FirstOrDefault();
    if (ln == null)
        throw new Exception("SCENARIO_FAIL: ln (Outline) not found in chartSpace spPr");

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
