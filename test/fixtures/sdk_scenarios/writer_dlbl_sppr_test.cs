// Validates that a writer-generated XLSX with a dLbl containing spPr is valid.
var validator = new OpenXmlValidator();
var errors = validator.Validate(SpreadsheetDocument.Open(XlsxPath, false));
foreach (var error in errors) Console.Error.WriteLine("VALIDATION_ERROR: " + error.Description);
if (!errors.Any())
{
    using var doc = SpreadsheetDocument.Open(XlsxPath, false);
    var chartPart = doc.WorkbookPart!.WorksheetParts.First()
                       .DrawingsPart!.ChartParts.First();
    var xml = chartPart.ChartSpace.OuterXml;
    if (!xml.Contains("AABB00") || !xml.Contains("112233"))
    {
        Console.Error.WriteLine("SCENARIO_FAIL: missing fill/line colors in dLbl spPr");
    }
    else
    {
        Console.Error.WriteLine("SCENARIO_PASS");
    }
}
