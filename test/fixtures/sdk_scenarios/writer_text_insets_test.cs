// Validates that an XLSX shape bodyPr has inset margin attributes.
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

    var bodyPr = txBody.Elements<DocumentFormat.OpenXml.Drawing.BodyProperties>().FirstOrDefault()
        ?? throw new Exception("BodyProperties is missing.");

    if (bodyPr.LeftInset?.Value != 91440)
        throw new Exception($"Expected lIns=91440, got {bodyPr.LeftInset?.Value}.");
    if (bodyPr.TopInset?.Value != 45720)
        throw new Exception($"Expected tIns=45720, got {bodyPr.TopInset?.Value}.");
    if (bodyPr.RightInset?.Value != 91440)
        throw new Exception($"Expected rIns=91440, got {bodyPr.RightInset?.Value}.");
    if (bodyPr.BottomInset?.Value != 45720)
        throw new Exception($"Expected bIns=45720, got {bodyPr.BottomInset?.Value}.");

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
