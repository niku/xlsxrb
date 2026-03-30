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
    var stylesPart = workbookPart.WorkbookStylesPart ?? throw new Exception("WorkbookStylesPart is missing.");
    var styleSheet = stylesPart.Stylesheet ?? throw new Exception("Stylesheet is missing.");

    var cellXfs = styleSheet.CellFormats?.Elements<CellFormat>().ToList()
        ?? throw new Exception("cellXfs are missing.");

    if (cellXfs.Count < 2)
        throw new Exception($"Expected at least 2 cellXfs but got {cellXfs.Count}.");

    var linkedXf = cellXfs[1];
    if (linkedXf.FormatId?.Value != 1U)
        throw new Exception($"Expected cellXf[1] xfId=1 but got {linkedXf.FormatId?.Value}.");
}
finally
{
    document.Dispose();
}
