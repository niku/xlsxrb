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
    var fonts = stylesPart.Stylesheet.Fonts;

    bool foundCharset = false;
    foreach (var font in fonts.Elements<Font>())
    {
        var charset = font.Elements<FontCharSet>().FirstOrDefault();
        if (charset != null && charset.Val?.Value == 128)
        {
            foundCharset = true;
            break;
        }
    }

    if (!foundCharset)
    {
        throw new Exception("No font found with charset=128 (Shift-JIS).");
    }
}
finally
{
    document.Dispose();
}
