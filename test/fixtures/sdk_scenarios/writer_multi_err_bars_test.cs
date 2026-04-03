// Validates that a writer-generated XLSX with multiple errBars (x and y directions) is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    var count = System.Text.RegularExpressions.Regex.Matches(xml, "<c:errBars>").Count;
    if (count != 2)
    {
        Console.Error.WriteLine($"SCENARIO_FAIL: expected 2 <c:errBars> elements but found {count}");
    }
    else if (!xml.Contains("<c:errDir val=\"x\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing x direction errBars");
    }
    else if (!xml.Contains("<c:errDir val=\"y\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing y direction errBars");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
