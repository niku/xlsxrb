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
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet definition is missing.");

    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!.Value!);
    var autoFilter = worksheetPart.Worksheet.GetFirstChild<AutoFilter>();
    if (autoFilter == null)
    {
        throw new Exception("AutoFilter element is missing.");
    }

    if (autoFilter.Reference?.Value != "A1:B10")
    {
        throw new Exception($"Expected AutoFilter ref 'A1:B10' but got '{autoFilter.Reference?.Value}'.");
    }
}
finally
{
    document.Dispose();
}
