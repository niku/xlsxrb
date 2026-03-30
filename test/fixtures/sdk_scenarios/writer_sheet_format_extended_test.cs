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
    var worksheetPart = workbookPart.WorksheetParts.First();
    var sfp = worksheetPart.Worksheet.Elements<SheetFormatProperties>().FirstOrDefault();
    if (sfp == null)
    {
        throw new Exception("SheetFormatProperties element is missing.");
    }

    if (sfp.OutlineLevelRow?.Value != 3)
    {
        throw new Exception($"Expected outlineLevelRow=3 but got '{sfp.OutlineLevelRow?.Value}'.");
    }

    if (sfp.OutlineLevelColumn?.Value != 2)
    {
        throw new Exception($"Expected outlineLevelCol=2 but got '{sfp.OutlineLevelColumn?.Value}'.");
    }

    if (sfp.ZeroHeight?.Value != true)
    {
        throw new Exception("Expected zeroHeight='true'.");
    }

    if (sfp.CustomHeight?.Value != true)
    {
        throw new Exception("Expected customHeight='true'.");
    }
}
finally
{
    document.Dispose();
}
