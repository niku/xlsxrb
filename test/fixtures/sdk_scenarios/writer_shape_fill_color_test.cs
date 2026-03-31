// Validates solidFill color on a shape via spPr.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var drawXml = wsPart.DrawingsPart.WorksheetDrawing.InnerXml;

    if (!drawXml.Contains("solidFill"))
        throw new Exception("SCENARIO_FAIL: solidFill not found in drawing XML");
    if (!drawXml.Contains("FF0000"))
        throw new Exception("SCENARIO_FAIL: fill color FF0000 not found in drawing XML");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
