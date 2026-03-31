// Validates fileRecoveryPr element in workbook XML.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var wbXml = wbPart.Workbook.InnerXml;

    if (!wbXml.Contains("fileRecoveryPr"))
        throw new Exception("SCENARIO_FAIL: fileRecoveryPr not found in workbook XML");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(5).ToList();
    if (errors.Count > 0)
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
