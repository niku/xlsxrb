var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var validationErrors = validator.Validate(document).Take(10).ToList();
    if (validationErrors.Any())
    {
        var message = string.Join(Environment.NewLine, validationErrors.Select(e => e.Description));
        throw new Exception($"OpenXmlValidator reported errors:{Environment.NewLine}{message}");
    }

    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var definedNames = workbookPart.Workbook.DefinedNames;
    if (definedNames == null)
    {
        throw new Exception("DefinedNames element is missing.");
    }

    var dnList = definedNames.Elements<DefinedName>().ToList();
    if (dnList.Count < 2)
    {
        throw new Exception($"Expected at least 2 defined names but got {dnList.Count}.");
    }

    if (dnList[0].Name?.Value != "MyRange")
    {
        throw new Exception($"Expected first defined name 'MyRange' but got '{dnList[0].Name?.Value}'.");
    }

    if (dnList[1].LocalSheetId?.Value != 1U)
    {
        throw new Exception($"Expected localSheetId=1 but got '{dnList[1].LocalSheetId?.Value}'.");
    }
}
finally
{
    document.Dispose();
}
