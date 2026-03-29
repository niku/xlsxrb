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
    var mergeCells = worksheetPart.Worksheet.GetFirstChild<MergeCells>();
    if (mergeCells == null)
    {
        throw new Exception("MergeCells element is missing.");
    }

    var mergeElements = mergeCells.Elements<MergeCell>().ToList();
    if (mergeElements.Count != 2)
    {
        throw new Exception($"Expected 2 mergeCell elements but got {mergeElements.Count}.");
    }

    if (mergeElements[0].Reference?.Value != "A1:B2")
    {
        throw new Exception($"Expected merge ref 'A1:B2' but got '{mergeElements[0].Reference?.Value}'.");
    }
    if (mergeElements[1].Reference?.Value != "C3:D4")
    {
        throw new Exception($"Expected merge ref 'C3:D4' but got '{mergeElements[1].Reference?.Value}'.");
    }
}
finally
{
    document.Dispose();
}
