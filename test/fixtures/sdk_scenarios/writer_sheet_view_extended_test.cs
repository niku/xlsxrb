var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var sheetViews = worksheetPart.Worksheet.SheetViews;
    if (sheetViews == null) throw new Exception("SheetViews element not found.");

    var sv = sheetViews.Elements<SheetView>().First();
    if (sv.ShowZeros == null || sv.ShowZeros.Value != false)
        throw new Exception($"Expected showZeros=false but got {sv.ShowZeros?.Value}.");
    if (sv.View == null || sv.View.Value != SheetViewValues.PageBreakPreview)
        throw new Exception($"Expected view=pageBreakPreview but got {sv.View?.Value}.");
    if (sv.ShowOutlineSymbols == null || sv.ShowOutlineSymbols.Value != false)
        throw new Exception($"Expected showOutlineSymbols=false but got {sv.ShowOutlineSymbols?.Value}.");
    if (sv.ShowRuler == null || sv.ShowRuler.Value != false)
        throw new Exception($"Expected showRuler=false but got {sv.ShowRuler?.Value}.");
}
finally
{
    document.Dispose();
}
