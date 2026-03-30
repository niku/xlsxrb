var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    // Find the pivot table part
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var pivotTableParts = worksheetPart.PivotTableParts.ToList();
    if (pivotTableParts.Count == 0) throw new Exception("No pivot table parts found.");

    var ptDef = pivotTableParts[0].PivotTableDefinition;
    var psi = ptDef.PivotTableStyle;
    if (psi == null) throw new Exception("PivotTableStyleInfo element not found.");
    if (psi.Name == null || psi.Name.Value != "PivotStyleLight16")
        throw new Exception($"Expected name=PivotStyleLight16 but got {psi.Name?.Value}.");
    if (psi.ShowRowHeaders == null || psi.ShowRowHeaders.Value != true)
        throw new Exception($"Expected showRowHeaders=true but got {psi.ShowRowHeaders?.Value}.");
    if (psi.ShowColumnHeaders == null || psi.ShowColumnHeaders.Value != true)
        throw new Exception($"Expected showColHeaders=true but got {psi.ShowColumnHeaders?.Value}.");
    if (psi.ShowRowStripes == null || psi.ShowRowStripes.Value != false)
        throw new Exception($"Expected showRowStripes=false but got {psi.ShowRowStripes?.Value}.");
    if (psi.ShowColumnStripes == null || psi.ShowColumnStripes.Value != false)
        throw new Exception($"Expected showColStripes=false but got {psi.ShowColumnStripes?.Value}.");
    if (psi.ShowLastColumn == null || psi.ShowLastColumn.Value != true)
        throw new Exception($"Expected showLastColumn=true but got {psi.ShowLastColumn?.Value}.");
}
finally
{
    document.Dispose();
}
