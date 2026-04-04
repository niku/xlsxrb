// Validates that a writer-generated XLSX with dLbl layout is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:dLbl"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:dLbl> element");
    }
    else if (!xml.Contains("<c:layout"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:layout> inside dLbl");
    }
    else if (!xml.Contains("<c:manualLayout"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:manualLayout> inside dLbl layout");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
