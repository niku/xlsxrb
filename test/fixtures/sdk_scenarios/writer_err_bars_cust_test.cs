// Validates that a writer-generated XLSX with custom error bars (plus/minus) is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:errBars"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:errBars> element");
    }
    else if (!xml.Contains("<c:plus>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:plus> element");
    }
    else if (!xml.Contains("<c:minus>"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:minus> element");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
