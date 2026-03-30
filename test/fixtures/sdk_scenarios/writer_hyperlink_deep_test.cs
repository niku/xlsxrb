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
    if (hlElements.Count != 3)
    {
        throw new Exception($"Expected 3 hyperlinks but got {hlElements.Count}.");
    }

    // HL1: A1 - external with display and tooltip
    var hl1 = hlElements[0];
    if (hl1.Reference?.Value != "A1")
        throw new Exception($"HL1 ref: expected 'A1' but got '{hl1.Reference?.Value}'.");
    if (hl1.Display?.Value != "Example Site")
        throw new Exception($"HL1 display: expected 'Example Site' but got '{hl1.Display?.Value}'.");
    if (hl1.Tooltip?.Value != "Click to visit")
        throw new Exception($"HL1 tooltip: expected 'Click to visit' but got '{hl1.Tooltip?.Value}'.");

    var rel1 = worksheetPart.HyperlinkRelationships
        .FirstOrDefault(r => r.Id == hl1.Id?.Value)
        ?? throw new Exception("HL1 hyperlink relationship is missing.");
    if (!rel1.Uri.ToString().TrimEnd('/').Equals("https://example.com"))
        throw new Exception($"HL1 URL: expected 'https://example.com' but got '{rel1.Uri}'.");

    // HL2: B1 - external with location (and url)
    var hl2 = hlElements[1];
    if (hl2.Reference?.Value != "B1")
        throw new Exception($"HL2 ref: expected 'B1' but got '{hl2.Reference?.Value}'.");
    if (hl2.Location?.Value != "Sheet2!A1")
        throw new Exception($"HL2 location: expected 'Sheet2!A1' but got '{hl2.Location?.Value}'.");

    var rel2 = worksheetPart.HyperlinkRelationships
        .FirstOrDefault(r => r.Id == hl2.Id?.Value)
        ?? throw new Exception("HL2 hyperlink relationship is missing.");
    if (!rel2.Uri.ToString().TrimEnd('/').Equals("https://example.com/page"))
        throw new Exception($"HL2 URL: expected 'https://example.com/page' but got '{rel2.Uri}'.");

    // HL3: C1 - internal only (location, no URL/r:id)
    var hl3 = hlElements[2];
    if (hl3.Reference?.Value != "C1")
        throw new Exception($"HL3 ref: expected 'C1' but got '{hl3.Reference?.Value}'.");
    if (hl3.Location?.Value != "Sheet1!D1")
        throw new Exception($"HL3 location: expected 'Sheet1!D1' but got '{hl3.Location?.Value}'.");
    if (hl3.Id?.Value != null)
        throw new Exception($"HL3 should have no r:id but got '{hl3.Id?.Value}'.");
}
finally
{
    document.Dispose();
}
