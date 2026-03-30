var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var stylesPart = document.WorkbookPart.WorkbookStylesPart;
    var stylesheet = stylesPart.Stylesheet;
    var tableStyles = stylesheet.TableStyles;
    if (tableStyles == null) throw new Exception("TableStyles element not found.");
    if (tableStyles.DefaultTableStyle == null || tableStyles.DefaultTableStyle.Value != "TableStyleMedium2")
        throw new Exception($"Expected defaultTableStyle=TableStyleMedium2 but got {tableStyles.DefaultTableStyle?.Value}.");
    if (tableStyles.DefaultPivotStyle == null || tableStyles.DefaultPivotStyle.Value != "PivotStyleLight16")
        throw new Exception($"Expected defaultPivotStyle=PivotStyleLight16 but got {tableStyles.DefaultPivotStyle?.Value}.");
}
finally
{
    document.Dispose();
}
