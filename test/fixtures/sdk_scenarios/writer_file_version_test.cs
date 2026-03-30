var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var wb = document.WorkbookPart.Workbook;
    var fv = wb.FileVersion;
    if (fv == null) throw new Exception("FileVersion element not found.");
    if (fv.ApplicationName == null || fv.ApplicationName.Value != "xl")
        throw new Exception($"Expected appName=xl but got {fv.ApplicationName?.Value}.");
    if (fv.LastEdited == null || fv.LastEdited.Value != "7")
        throw new Exception($"Expected lastEdited=7 but got {fv.LastEdited?.Value}.");
    if (fv.LowestEdited == null || fv.LowestEdited.Value != "7")
        throw new Exception($"Expected lowestEdited=7 but got {fv.LowestEdited?.Value}.");
    if (fv.BuildVersion == null || fv.BuildVersion.Value != "27425")
        throw new Exception($"Expected rupBuild=27425 but got {fv.BuildVersion?.Value}.");
}
finally
{
    document.Dispose();
}
