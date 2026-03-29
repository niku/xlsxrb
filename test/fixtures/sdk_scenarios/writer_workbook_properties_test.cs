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
    var workbook = workbookPart.Workbook;

    var wbPr = workbook.WorkbookProperties;
    if (wbPr == null)
    {
        throw new Exception("WorkbookProperties element is missing.");
    }

    var bookViews = workbook.BookViews;
    if (bookViews == null)
    {
        throw new Exception("BookViews element is missing.");
    }

    var wbView = bookViews.GetFirstChild<WorkbookView>();
    if (wbView == null)
    {
        throw new Exception("WorkbookView element is missing.");
    }

    if (wbView.ActiveTab?.Value != 1U)
    {
        throw new Exception($"Expected activeTab=1 but got '{wbView.ActiveTab?.Value}'.");
    }

    var calcPr = workbook.CalculationProperties;
    if (calcPr == null)
    {
        throw new Exception("CalculationProperties element is missing.");
    }
}
finally
{
    document.Dispose();
}
