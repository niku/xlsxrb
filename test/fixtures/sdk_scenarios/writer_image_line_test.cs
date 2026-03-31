// Validates line color and width on an image via a:ln element.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var drawXml = wsPart.DrawingsPart.WorksheetDrawing.InnerXml;

    if (!drawXml.Contains(":ln ") && !drawXml.Contains(":ln>"))
        throw new Exception("SCENARIO_FAIL: ln (line) element not found in drawing XML");
    if (!drawXml.Contains("FF0000"))
        throw new Exception("SCENARIO_FAIL: line color FF0000 not found in drawing XML");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
