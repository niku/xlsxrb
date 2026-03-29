// Validates sheet protection in XLSX.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("No sheets found.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
    var sp = worksheetPart.Worksheet.GetFirstChild<SheetProtection>()
        ?? throw new Exception("SheetProtection element is missing.");

    if (sp.Sheet == null || !sp.Sheet.Value)
        throw new Exception("sheet attribute should be true.");
    if (sp.Password == null || sp.Password.Value != "CF1A")
        throw new Exception($"Expected password CF1A but got {sp.Password?.Value}.");
}
finally
{
    document.Dispose();
}
