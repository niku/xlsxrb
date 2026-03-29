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
    var stylesPart = workbookPart.WorkbookStylesPart
        ?? throw new Exception("WorkbookStylesPart is missing.");

    var stylesheet = stylesPart.Stylesheet
        ?? throw new Exception("Stylesheet is missing.");

    // Verify numFmts
    var numFmts = stylesheet.NumberingFormats;
    if (numFmts == null || numFmts.Count?.Value != 1)
    {
        throw new Exception($"Expected 1 numFmt but got {numFmts?.Count?.Value ?? 0}.");
    }

    var numFmt = numFmts.Elements<NumberingFormat>().First();
    if (numFmt.NumberFormatId?.Value != 164)
    {
        throw new Exception($"Expected numFmtId 164 but got {numFmt.NumberFormatId?.Value}.");
    }
    if (numFmt.FormatCode?.Value != "0.00")
    {
        throw new Exception($"Expected formatCode '0.00' but got '{numFmt.FormatCode?.Value}'.");
    }

    // Verify cellXfs has 2 entries (default + custom)
    var cellXfs = stylesheet.CellFormats;
    if (cellXfs == null || cellXfs.Count?.Value != 2)
    {
        throw new Exception($"Expected 2 cellXfs but got {cellXfs?.Count?.Value ?? 0}.");
    }

    // Verify cell A1 uses style index 1
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!.Value!);
    var a1Cell = worksheetPart.Worksheet.Descendants<Cell>()
        .FirstOrDefault(c => c.CellReference?.Value == "A1")
        ?? throw new Exception("A1 cell is missing.");

    if (a1Cell.StyleIndex?.Value != 1)
    {
        throw new Exception($"Expected A1 styleIndex=1 but got {a1Cell.StyleIndex?.Value}.");
    }
}
finally
{
    document.Dispose();
}
