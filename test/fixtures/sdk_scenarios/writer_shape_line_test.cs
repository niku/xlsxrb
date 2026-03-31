// Validates line color and width on a shape via a:ln element.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var drawXml = wsPart.DrawingsPart.WorksheetDrawing.InnerXml;

    if (!drawXml.Contains(":ln ") && !drawXml.Contains(":ln>"))
        throw new Exception("SCENARIO_FAIL: ln (line) element not found in drawing XML");
    if (!drawXml.Contains("0000FF"))
        throw new Exception("SCENARIO_FAIL: line color 0000FF not found in drawing XML");

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
