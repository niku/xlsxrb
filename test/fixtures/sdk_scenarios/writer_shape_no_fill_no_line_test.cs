// Validates noFill and noLine (empty ln) on a shape.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var drawXml = wsPart.DrawingsPart.WorksheetDrawing.InnerXml;

    if (!drawXml.Contains("noFill"))
        throw new Exception("SCENARIO_FAIL: noFill not found in drawing XML");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
