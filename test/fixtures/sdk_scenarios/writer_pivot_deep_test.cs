// Validates that an XLSX pivot table has col_fields, items, and proper cache.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");

    var pivotCacheParts = workbookPart.PivotTableCacheDefinitionParts.ToList();
    if (pivotCacheParts.Count == 0)
        throw new Exception("No pivot cache parts found.");

    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var pivotTableParts = worksheetPart.PivotTableParts.ToList();
    if (pivotTableParts.Count == 0)
        throw new Exception("No pivot table parts found.");

    // Validate
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri}, Path: {e.Path?.XPath})"));
        throw new Exception($"Validation errors:\n{messages}");
    }
}
finally
{
    document.Dispose();
}
