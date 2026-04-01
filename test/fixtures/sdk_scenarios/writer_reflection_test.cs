// Validates that an XLSX shape has a:effectLst with a:reflection element.
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

    var spPr = sp.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties>().FirstOrDefault()
        ?? throw new Exception("ShapeProperties is missing.");

    var effectLst = spPr.Elements<DocumentFormat.OpenXml.Drawing.EffectList>().FirstOrDefault()
        ?? throw new Exception("EffectList is missing.");

    var reflection = effectLst.Elements<DocumentFormat.OpenXml.Drawing.Reflection>().FirstOrDefault()
        ?? throw new Exception("Reflection is missing.");

    if (reflection.BlurRadius?.Value != 6350L) throw new Exception($"Expected blurRad=6350, got {reflection.BlurRadius?.Value}.");
    if (reflection.StartOpacity?.Value != 52000) throw new Exception($"Expected stA=52000, got {reflection.StartOpacity?.Value}.");
    if (reflection.EndAlpha?.Value != 300) throw new Exception($"Expected endA=300, got {reflection.EndAlpha?.Value}.");
    if (reflection.Direction?.Value != 5400000) throw new Exception($"Expected dir=5400000, got {reflection.Direction?.Value}.");
    if (reflection.VerticalRatio?.Value != -100000) throw new Exception($"Expected sy=-100000, got {reflection.VerticalRatio?.Value}.");
    if (reflection.Alignment?.Value != DocumentFormat.OpenXml.Drawing.RectangleAlignmentValues.BottomLeft)
        throw new Exception($"Expected algn=bl, got {reflection.Alignment?.Value}.");

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
