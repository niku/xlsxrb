var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var stylesPart = document.WorkbookPart.WorkbookStylesPart;
    var stylesheet = stylesPart.Stylesheet;
    var colors = stylesheet.Colors;
    if (colors == null) throw new Exception("Colors element not found.");
    var indexedColors = colors.IndexedColors;
    if (indexedColors == null) throw new Exception("IndexedColors element not found.");
    var rgbList = indexedColors.Elements<RgbColor>().ToList();
    if (rgbList.Count < 3) throw new Exception($"Expected at least 3 indexed colors but got {rgbList.Count}.");
    if (rgbList[0].Rgb == null || rgbList[0].Rgb.Value != "FF000000")
        throw new Exception($"Expected first color=FF000000 but got {rgbList[0].Rgb?.Value}.");
    if (rgbList[1].Rgb == null || rgbList[1].Rgb.Value != "FFFFFFFF")
        throw new Exception($"Expected second color=FFFFFFFF but got {rgbList[1].Rgb?.Value}.");
    if (rgbList[2].Rgb == null || rgbList[2].Rgb.Value != "FFFF0000")
        throw new Exception($"Expected third color=FFFF0000 but got {rgbList[2].Rgb?.Value}.");
}
finally
{
    document.Dispose();
}
