// Validates that an XLSX shape paragraph has spcBef and spcAft spacing elements.
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

    var sp = anchors[0].Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault()
        ?? throw new Exception("Shape is missing.");

    var txBody = sp.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody>().FirstOrDefault()
        ?? throw new Exception("TextBody is missing.");

    var para = txBody.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault()
        ?? throw new Exception("Paragraph is missing.");

    var pPr = para.ParagraphProperties ?? throw new Exception("ParagraphProperties is missing.");

    var spcBef = pPr.SpaceBefore ?? throw new Exception("SpaceBefore is missing.");
    var spcBefPts = spcBef.Elements<DocumentFormat.OpenXml.Drawing.SpacingPoints>().FirstOrDefault()
        ?? throw new Exception("SpaceBefore SpacingPoints is missing.");
    if (spcBefPts.Val?.Value != 600)
        throw new Exception($"Expected spcBef=600, got {spcBefPts.Val?.Value}.");

    var spcAft = pPr.SpaceAfter ?? throw new Exception("SpaceAfter is missing.");
    var spcAftPts = spcAft.Elements<DocumentFormat.OpenXml.Drawing.SpacingPoints>().FirstOrDefault()
        ?? throw new Exception("SpaceAfter SpacingPoints is missing.");
    if (spcAftPts.Val?.Value != 400)
        throw new Exception($"Expected spcAft=400, got {spcAftPts.Val?.Value}.");

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
