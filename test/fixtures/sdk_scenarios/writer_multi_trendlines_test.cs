// Validates that a writer-generated XLSX with multiple trendlines is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    int count = 0;
    int idx = 0;
    while ((idx = xml.IndexOf("<c:trendline>", idx)) >= 0)
    {
        count++;
        idx += 13;
    }
    if (count < 2)
    {
        Console.Error.WriteLine($"SCENARIO_FAIL: expected at least 2 trendlines, found {count}");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
