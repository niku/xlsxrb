var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var ws = worksheetPart.Worksheet;

    // CellWatches is in the x: namespace
    var cellWatches = ws.GetFirstChild<CellWatches>();
    if (cellWatches == null) throw new Exception("CellWatches element not found.");

    var watches = cellWatches.Elements<CellWatch>().ToList();
    if (watches.Count != 2) throw new Exception($"Expected 2 cell watches but got {watches.Count}.");
    if (watches[0].CellReference == null || watches[0].CellReference.Value != "A1")
        throw new Exception($"Expected first watch r=A1 but got {watches[0].CellReference?.Value}.");
    if (watches[1].CellReference == null || watches[1].CellReference.Value != "B2")
        throw new Exception($"Expected second watch r=B2 but got {watches[1].CellReference?.Value}.");
}
finally
{
    document.Dispose();
}
