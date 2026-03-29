// Validates shared and array formulas in XLSX.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart missing.");
    var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("No sheets.");
    var wsPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
    var cells = wsPart.Worksheet.Descendants<Cell>().ToList();

    var b1 = cells.FirstOrDefault(c => c.CellReference == "B1")
        ?? throw new Exception("B1 missing.");
    var b1f = b1.CellFormula ?? throw new Exception("B1 formula missing.");
    if (b1f.FormulaType?.Value != CellFormulaValues.Shared)
        throw new Exception($"B1 should be shared, got {b1f.FormulaType?.Value}.");
    if (b1f.Reference?.Value != "B1:B2")
        throw new Exception($"B1 ref should be B1:B2, got {b1f.Reference?.Value}.");

    var c1 = cells.FirstOrDefault(c => c.CellReference == "C1")
        ?? throw new Exception("C1 missing.");
    var c1f = c1.CellFormula ?? throw new Exception("C1 formula missing.");
    if (c1f.FormulaType?.Value != CellFormulaValues.Array)
        throw new Exception($"C1 should be array, got {c1f.FormulaType?.Value}.");
}
finally
{
    document.Dispose();
}
