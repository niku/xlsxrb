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
    var hyperlinks = worksheetPart.Worksheet.GetFirstChild<Hyperlinks>();
    if (hyperlinks == null)
    {
        throw new Exception("Hyperlinks element is missing.");
    }

    var hlElements = hyperlinks.Elements<Hyperlink>().ToList();
    if (hlElements.Count != 1)
    {
        throw new Exception($"Expected 1 hyperlink but got {hlElements.Count}.");
    }

    var hl = hlElements[0];
    if (hl.Reference?.Value != "A1")
    {
        throw new Exception($"Expected hyperlink ref 'A1' but got '{hl.Reference?.Value}'.");
    }

    // Verify the relationship target
    var hyperlinkRel = worksheetPart.HyperlinkRelationships
        .FirstOrDefault(r => r.Id == hl.Id?.Value)
        ?? throw new Exception("Hyperlink relationship is missing.");

    if (!hyperlinkRel.Uri.ToString().TrimEnd('/').Equals("https://example.com"))
    {
        throw new Exception($"Expected URL 'https://example.com' but got '{hyperlinkRel.Uri}'.");
    }
}
finally
{
    document.Dispose();
}
