// Validates that an XLSX shape has a:effectLst with a:outerShdw element.
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

    var outerShdw = effectLst.Elements<DocumentFormat.OpenXml.Drawing.OuterShadow>().FirstOrDefault()
        ?? throw new Exception("OuterShadow is missing.");

    if (outerShdw.BlurRadius?.Value != 50800L) throw new Exception($"Expected blurRad=50800, got {outerShdw.BlurRadius?.Value}.");
    if (outerShdw.Distance?.Value != 38100L) throw new Exception($"Expected dist=38100, got {outerShdw.Distance?.Value}.");
    if (outerShdw.Direction?.Value != 2700000) throw new Exception($"Expected dir=2700000, got {outerShdw.Direction?.Value}.");

    var srgbClr = outerShdw.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()
        ?? throw new Exception("RgbColorModelHex is missing.");
    if (srgbClr.Val?.Value != "000000") throw new Exception($"Expected color 000000, got {srgbClr.Val?.Value}.");

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
