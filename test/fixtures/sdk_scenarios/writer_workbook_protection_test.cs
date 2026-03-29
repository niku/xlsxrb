// Validates workbook protection in XLSX.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var wp = workbookPart.Workbook.GetFirstChild<WorkbookProtection>()
        ?? throw new Exception("WorkbookProtection element is missing.");

    if (wp.LockStructure == null || !wp.LockStructure.Value)
        throw new Exception("lockStructure should be true.");
}
finally
{
    document.Dispose();
}
