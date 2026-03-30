var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var wb = document.WorkbookPart.Workbook;

    // The conformance attribute may not be exposed as a strongly-typed property in all SDK versions.
    // Check the raw XML to verify the attribute is present.
    var outerXml = wb.OuterXml;
    if (!outerXml.Contains("conformance=\"transitional\""))
        throw new Exception($"Expected conformance=transitional in workbook XML but not found.\nXML: {outerXml.Substring(0, Math.Min(500, outerXml.Length))}");
}
finally
{
    document.Dispose();
}
