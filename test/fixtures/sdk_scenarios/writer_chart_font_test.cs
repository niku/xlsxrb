// Validates that a writer-generated XLSX with chart-level txPr is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    // txPr should appear as a direct child of chartSpace (not inside chart element)
    if (!xml.Contains("<c:txPr"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing <c:txPr> element in chartSpace");
    }
    else if (!xml.Contains("typeface=\"Arial\""))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing Arial typeface in txPr");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
