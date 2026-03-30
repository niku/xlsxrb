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
    var sheetProtection = worksheetPart.Worksheet.Elements<SheetProtection>().FirstOrDefault();
    if (sheetProtection == null)
    {
        throw new Exception("SheetProtection element is missing.");
    }

    if (sheetProtection.AlgorithmName?.Value != "SHA-512")
    {
        throw new Exception($"Expected algorithmName='SHA-512' but got '{sheetProtection.AlgorithmName?.Value}'.");
    }

    if (string.IsNullOrEmpty(sheetProtection.HashValue?.Value))
    {
        throw new Exception("hashValue is missing or empty.");
    }

    if (string.IsNullOrEmpty(sheetProtection.SaltValue?.Value))
    {
        throw new Exception("saltValue is missing or empty.");
    }

    if (sheetProtection.SpinCount?.Value != 1000U)
    {
        throw new Exception($"Expected spinCount=1000 but got '{sheetProtection.SpinCount?.Value}'.");
    }

    if (sheetProtection.Sheet?.Value != true)
    {
        throw new Exception("Expected sheet='1' (true).");
    }
}
finally
{
    document.Dispose();
}
