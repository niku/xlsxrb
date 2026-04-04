// Validates that a writer-generated XLSX with chart protection is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:protection"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:protection> element");
    }
    else if (!xml.Contains("<c:chartObject"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:chartObject> inside protection");
    }
    else if (!xml.Contains("<c:data"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:data> inside protection");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
