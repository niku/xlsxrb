var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var coreProps = document.PackageProperties;
    if (coreProps.Title != "My Workbook")
    {
        throw new Exception($"Expected title 'My Workbook' but got '{coreProps.Title}'.");
    }
    if (coreProps.Creator != "Test User")
    {
        throw new Exception($"Expected creator 'Test User' but got '{coreProps.Creator}'.");
    }
    if (coreProps.Created == null)
    {
        throw new Exception("Expected created date but got null.");
    }
    if (coreProps.Modified == null)
    {
        throw new Exception("Expected modified date but got null.");
    }
}
finally
{
    document.Dispose();
}
