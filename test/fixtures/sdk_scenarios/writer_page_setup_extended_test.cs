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
    var pageSetup = worksheetPart.Worksheet.Elements<PageSetup>().FirstOrDefault();
    if (pageSetup == null)
    {
        throw new Exception("PageSetup element is missing.");
    }

    if (pageSetup.PageOrder?.Value != PageOrderValues.OverThenDown)
    {
        throw new Exception($"Expected pageOrder='overThenDown' but got '{pageSetup.PageOrder?.Value}'.");
    }

    if (pageSetup.BlackAndWhite?.Value != true)
    {
        throw new Exception("Expected blackAndWhite='true'.");
    }

    if (pageSetup.Draft?.Value != true)
    {
        throw new Exception("Expected draft='true'.");
    }

    if (pageSetup.CellComments?.Value != CellCommentsValues.AtEnd)
    {
        throw new Exception($"Expected cellComments='atEnd' but got '{pageSetup.CellComments?.Value}'.");
    }

    if (pageSetup.FirstPageNumber?.Value != 5U)
    {
        throw new Exception($"Expected firstPageNumber=5 but got '{pageSetup.FirstPageNumber?.Value}'.");
    }

    if (pageSetup.UseFirstPageNumber?.Value != true)
    {
        throw new Exception("Expected useFirstPageNumber='true'.");
    }

    if (pageSetup.HorizontalDpi?.Value != 300U)
    {
        throw new Exception($"Expected horizontalDpi=300 but got '{pageSetup.HorizontalDpi?.Value}'.");
    }

    if (pageSetup.VerticalDpi?.Value != 300U)
    {
        throw new Exception($"Expected verticalDpi=300 but got '{pageSetup.VerticalDpi?.Value}'.");
    }
}
finally
{
    document.Dispose();
}
