// Validates XLSX with rich text in shared strings.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart missing.");
    var ssPart = workbookPart.SharedStringTablePart ?? throw new Exception("SharedStringTablePart missing.");
    var items = ssPart.SharedStringTable.Elements<SharedStringItem>().ToList();
    if (items.Count == 0) throw new Exception("No shared string items.");

    var runs = items[0].Elements<Run>().ToList();
    if (runs.Count < 2) throw new Exception($"Expected >=2 runs, got {runs.Count}.");

    var firstRunText = runs[0].GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Text>()?.Text;
    if (firstRunText != "Bold")
        throw new Exception($"Expected first run text 'Bold' but got '{firstRunText}'.");

    var rpr = runs[0].GetFirstChild<RunProperties>();
    if (rpr?.GetFirstChild<Bold>() == null)
        throw new Exception("First run should be bold.");
}
finally
{
    document.Dispose();
}
