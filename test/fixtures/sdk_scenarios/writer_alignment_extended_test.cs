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
    var stylesPart = workbookPart.WorkbookStylesPart ?? throw new Exception("WorkbookStylesPart is missing.");
    var stylesheet = stylesPart.Stylesheet;

    // Find a cellXf that has alignment with readingOrder
    var cellXfs = stylesheet.CellFormats;
    bool foundReadingOrder = false;
    bool foundJustifyLastLine = false;

    foreach (var xf in cellXfs.Elements<CellFormat>())
    {
        var alignment = xf.Elements<Alignment>().FirstOrDefault();
        if (alignment != null)
        {
            if (alignment.ReadingOrder?.Value == 2U)
            {
                foundReadingOrder = true;
            }
            if (alignment.JustifyLastLine?.Value == true)
            {
                foundJustifyLastLine = true;
            }
        }
    }

    if (!foundReadingOrder)
    {
        throw new Exception("No cellXf found with readingOrder=2.");
    }

    if (!foundJustifyLastLine)
    {
        throw new Exception("No cellXf found with justifyLastLine=true.");
    }
}
finally
{
    document.Dispose();
}
