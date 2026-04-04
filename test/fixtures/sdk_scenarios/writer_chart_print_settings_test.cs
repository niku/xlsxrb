// Validates that a writer-generated XLSX with chart printSettings is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:printSettings"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:printSettings> element");
    }
    else if (!xml.Contains("<c:headerFooter"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:headerFooter> inside printSettings");
    }
    else if (!xml.Contains("<c:pageMargins"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:pageMargins> inside printSettings");
    }
    else if (!xml.Contains("<c:pageSetup"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:pageSetup> inside printSettings");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
