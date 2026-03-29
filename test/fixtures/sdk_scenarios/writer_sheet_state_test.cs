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
    var sheetList = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
        ?? throw new Exception("Sheets are missing.");

    if (sheetList.Count != 3)
    {
        throw new Exception($"Expected 3 sheets but got {sheetList.Count}.");
    }

    if (sheetList[0].State?.Value != null && sheetList[0].State?.Value != SheetStateValues.Visible)
    {
        throw new Exception($"Expected Sheet1 visible but got '{sheetList[0].State?.Value}'.");
    }

    if (sheetList[1].State?.Value != SheetStateValues.Hidden)
    {
        throw new Exception($"Expected Hidden sheet hidden but got '{sheetList[1].State?.Value}'.");
    }

    if (sheetList[2].State?.Value != SheetStateValues.VeryHidden)
    {
        throw new Exception($"Expected VeryHidden sheet veryHidden but got '{sheetList[2].State?.Value}'.");
    }
}
finally
{
    document.Dispose();
}
