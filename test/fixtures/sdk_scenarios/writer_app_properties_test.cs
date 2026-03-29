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

    var props = document.ExtendedFilePropertiesPart?.Properties;
    if (props == null)
    {
        throw new Exception("ExtendedFilePropertiesPart is missing.");
    }

    if (props.Application?.Text != "Xlsxrb")
    {
        throw new Exception($"Expected Application 'Xlsxrb' but got '{props.Application?.Text}'.");
    }

    if (props.ApplicationVersion?.Text != "1.0.0")
    {
        throw new Exception($"Expected AppVersion '1.0.0' but got '{props.ApplicationVersion?.Text}'.");
    }
}
finally
{
    document.Dispose();
}
