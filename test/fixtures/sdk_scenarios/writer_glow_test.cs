// Validates that an XLSX shape has a:effectLst with a:glow element.
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

    var glow = effectLst.Elements<DocumentFormat.OpenXml.Drawing.Glow>().FirstOrDefault()
        ?? throw new Exception("Glow is missing.");

    if (glow.Radius?.Value != 101600L) throw new Exception($"Expected rad=101600, got {glow.Radius?.Value}.");

    var srgbClr = glow.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()
        ?? throw new Exception("RgbColorModelHex is missing.");
    if (srgbClr.Val?.Value != "FF0000") throw new Exception($"Expected color FF0000, got {srgbClr.Val?.Value}.");

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
