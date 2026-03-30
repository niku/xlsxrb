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
    var worksheetPart = workbookPart.WorksheetParts.First();
    var dvs = worksheetPart.Worksheet.Elements<DataValidations>().FirstOrDefault();
    if (dvs == null)
    {
        throw new Exception("DataValidations element is missing.");
    }

    var dv = dvs.Elements<DataValidation>().First();

    if (dv.ShowDropDown?.Value != true)
    {
        throw new Exception($"Expected showDropDown=true but got '{dv.ShowDropDown?.Value}'.");
    }

    if (dv.ImeMode?.Value != DataValidationImeModeValues.Hiragana)
    {
        throw new Exception($"Expected imeMode='hiragana' but got '{dv.ImeMode?.Value}'.");
    }

    if (dv.Type?.Value != DataValidationValues.List)
    {
        throw new Exception($"Expected type='list' but got '{dv.Type?.Value}'.");
    }
}
finally
{
    document.Dispose();
}
