var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var wb = document.WorkbookPart.Workbook;
    var fs = wb.FileSharing;
    if (fs == null) throw new Exception("FileSharing element not found.");
    if (fs.ReadOnlyRecommended == null || fs.ReadOnlyRecommended.Value != true)
        throw new Exception($"Expected readOnlyRecommended=true but got {fs.ReadOnlyRecommended?.Value}.");
    if (fs.UserName == null || fs.UserName.Value != "TestUser")
        throw new Exception($"Expected userName=TestUser but got {fs.UserName?.Value}.");
}
finally
{
    document.Dispose();
}
