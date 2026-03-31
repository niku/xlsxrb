// Validates that an XLSX image contains a:srcRect for cropping inside blipFill.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    var wsDr = drawingsPart.WorksheetDrawing
        ?? throw new Exception("WorksheetDrawing is missing.");

    var anchors = wsDr.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>().ToList();
    if (anchors.Count == 0) throw new Exception("No anchors found.");

    var pic = anchors[0].Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().FirstOrDefault()
        ?? throw new Exception("Picture is missing.");

    var blipFill = pic.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill>().FirstOrDefault()
        ?? throw new Exception("BlipFill is missing.");

    var srcRect = blipFill.Elements<DocumentFormat.OpenXml.Drawing.SourceRectangle>().FirstOrDefault()
        ?? throw new Exception("SourceRectangle is missing.");

    if (srcRect.Top?.Value != 10000) throw new Exception($"Expected top=10000, got {srcRect.Top?.Value}.");
    if (srcRect.Bottom?.Value != 20000) throw new Exception($"Expected bottom=20000, got {srcRect.Bottom?.Value}.");
    if (srcRect.Left?.Value != 5000) throw new Exception($"Expected left=5000, got {srcRect.Left?.Value}.");
    if (srcRect.Right?.Value != 15000) throw new Exception($"Expected right=15000, got {srcRect.Right?.Value}.");

    // Validate
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri}, Path: {e.Path?.XPath})"));
        throw new Exception($"Validation errors:\n{messages}");
    }
}
finally
{
    document.Dispose();
}
