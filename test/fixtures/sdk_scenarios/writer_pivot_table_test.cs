// Validates that an XLSX contains a pivot table definition.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheets = workbookPart.Workbook.Sheets?.Elements<Sheet>().ToList()
        ?? throw new Exception("No sheets found.");

    // Find the sheet with a pivot table (second sheet).
    if (sheets.Count < 2)
        throw new Exception("Expected at least 2 sheets.");

    var pivotSheet = (WorksheetPart)workbookPart.GetPartById(sheets[1].Id.Value);
    var pivotParts = pivotSheet.PivotTableParts.ToList();
    if (pivotParts.Count == 0)
        throw new Exception("No pivot table parts found.");

    var ptDef = pivotParts[0].PivotTableDefinition
        ?? throw new Exception("PivotTableDefinition is missing.");

    if (ptDef.Name?.Value == null)
        throw new Exception("Pivot table name is missing.");

    var dataFields = ptDef.Elements<DataFields>().FirstOrDefault();
    if (dataFields == null)
        throw new Exception("DataFields element is missing.");

    var firstDF = dataFields.Elements<DataField>().FirstOrDefault()
        ?? throw new Exception("No DataField found.");

    if (firstDF.Subtotal?.Value != DataConsolidateFunctionValues.Sum)
        throw new Exception($"Expected subtotal 'sum' but got '{firstDF.Subtotal?.Value}'.");
}
finally
{
    document.Dispose();
}
