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
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id!.Value!);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    // Find the pic element's cNvPr
    var wsDr = drawingsPart.WorksheetDrawing;
    var anchors = wsDr.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>().ToList();
    if (anchors.Count != 1)
        throw new Exception($"Expected 1 anchor but got {anchors.Count}.");

    var pic = anchors[0].Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().FirstOrDefault()
        ?? throw new Exception("Picture element is missing.");
    var nvPicPr = pic.GetFirstChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties>()
        ?? throw new Exception("nvPicPr is missing.");
    var cNvPr = nvPicPr.GetFirstChild<DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties>()
        ?? throw new Exception("cNvPr is missing.");

    if (cNvPr.Name?.Value != "Logo")
        throw new Exception($"Expected name 'Logo' but got '{cNvPr.Name?.Value}'.");
    if (cNvPr.Description?.Value != "Company logo image")
        throw new Exception($"Expected descr 'Company logo image' but got '{cNvPr.Description?.Value}'.");
}
finally
{
    document.Dispose();
}
