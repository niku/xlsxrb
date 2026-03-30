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
    var headerFooter = worksheetPart.Worksheet.Elements<HeaderFooter>().FirstOrDefault();
    if (headerFooter == null)
    {
        throw new Exception("HeaderFooter element is missing.");
    }

    if (headerFooter.DifferentFirst?.Value != true)
    {
        throw new Exception("Expected differentFirst='true'.");
    }

    if (headerFooter.DifferentOddEven?.Value != true)
    {
        throw new Exception("Expected differentOddEven='true'.");
    }

    var firstHeader = headerFooter.FirstHeader?.Text;
    if (firstHeader != "&CFirst Page Header")
    {
        throw new Exception($"Expected firstHeader='&CFirst Page Header' but got '{firstHeader}'.");
    }

    var firstFooter = headerFooter.FirstFooter?.Text;
    if (firstFooter != "&CFirst Page Footer")
    {
        throw new Exception($"Expected firstFooter='&CFirst Page Footer' but got '{firstFooter}'.");
    }

    var oddHeader = headerFooter.OddHeader?.Text;
    if (oddHeader != "&LOdd Header")
    {
        throw new Exception($"Expected oddHeader='&LOdd Header' but got '{oddHeader}'.");
    }

    var evenHeader = headerFooter.EvenHeader?.Text;
    if (evenHeader != "&LEven Header")
    {
        throw new Exception($"Expected evenHeader='&LEven Header' but got '{evenHeader}'.");
    }
}
finally
{
    document.Dispose();
}
