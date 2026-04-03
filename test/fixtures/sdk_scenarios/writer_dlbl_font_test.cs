// Validates that a writer-generated XLSX with a dLbl containing txPr (font) is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:txPr>") || !xml.Contains("FF0000"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:txPr> with font color in dLbl");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
