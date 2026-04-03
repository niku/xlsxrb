// Validates that a writer-generated XLSX with surface3DChart has 3 axis IDs and serAx.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:surface3DChart>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:surface3DChart> element");
    }
    else if (!xml.Contains("<c:serAx>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:serAx> element");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
