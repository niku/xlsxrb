// Validates that a writer-generated XLSX with per-series shape on bar3D chart is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("<c:shape val=\"cone\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing per-series <c:shape val=\"cone\"> element");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
