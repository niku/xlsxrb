var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var stylesPart = document.WorkbookPart.WorkbookStylesPart;
    var fonts = stylesPart.Stylesheet.Fonts;

    bool foundShadow = false;
    bool foundOutline = false;
    bool foundCondense = false;
    bool foundExtend = false;

    foreach (var font in fonts.Elements<Font>())
    {
        if (font.Shadow != null) foundShadow = true;
        if (font.Outline != null) foundOutline = true;
        if (font.Condense != null) foundCondense = true;
        if (font.Extend != null) foundExtend = true;
    }

    if (!foundShadow) throw new Exception("No font with shadow element found.");
    if (!foundOutline) throw new Exception("No font with outline element found.");
    if (!foundCondense) throw new Exception("No font with condense element found.");
    if (!foundExtend) throw new Exception("No font with extend element found.");
}
finally
{
    document.Dispose();
}
