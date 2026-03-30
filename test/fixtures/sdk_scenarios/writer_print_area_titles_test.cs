var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var validationErrors = validator.Validate(document).Take(10).ToList();
    if (validationErrors.Any())
    {
        var message = string.Join(Environment.NewLine, validationErrors.Select(e => e.Description));
        throw new Exception($"OpenXmlValidator reported errors:{Environment.NewLine}{message}");
    }

    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var definedNames = workbookPart.Workbook.DefinedNames;
    if (definedNames == null)
    {
        throw new Exception("DefinedNames element is missing.");
    }

    var dnList = definedNames.Elements<DefinedName>().ToList();
    if (dnList.Count < 2)
    {
        throw new Exception($"Expected at least 2 defined names but got {dnList.Count}.");
    }

    // Check _xlnm.Print_Area
    var printArea = dnList.FirstOrDefault(dn => dn.Name?.Value == "_xlnm.Print_Area");
    if (printArea == null)
    {
        throw new Exception("_xlnm.Print_Area defined name not found.");
    }
    var paValue = printArea.InnerText;
    if (!paValue.Contains("$A$1:$D$20"))
    {
        throw new Exception($"Expected Print_Area to contain '$A$1:$D$20' but got '{paValue}'.");
    }
    if (printArea.LocalSheetId?.Value != 0U)
    {
        throw new Exception($"Expected Print_Area localSheetId=0 but got '{printArea.LocalSheetId?.Value}'.");
    }

    // Check _xlnm.Print_Titles
    var printTitles = dnList.FirstOrDefault(dn => dn.Name?.Value == "_xlnm.Print_Titles");
    if (printTitles == null)
    {
        throw new Exception("_xlnm.Print_Titles defined name not found.");
    }
    var ptValue = printTitles.InnerText;
    if (!ptValue.Contains("$A:$B"))
    {
        throw new Exception($"Expected Print_Titles to contain '$A:$B' but got '{ptValue}'.");
    }
    if (!ptValue.Contains("$1:$3"))
    {
        throw new Exception($"Expected Print_Titles to contain '$1:$3' but got '{ptValue}'.");
    }
}
finally
{
    document.Dispose();
}
