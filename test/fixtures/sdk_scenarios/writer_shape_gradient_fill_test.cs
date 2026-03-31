// Validates that an XLSX shape has a:gradFill with gradient stops and linear direction.
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

    var gradFill = spPr.Elements<DocumentFormat.OpenXml.Drawing.GradientFill>().FirstOrDefault()
        ?? throw new Exception("GradientFill is missing.");

    var gsLst = gradFill.Elements<DocumentFormat.OpenXml.Drawing.GradientStopList>().FirstOrDefault()
        ?? throw new Exception("GradientStopList is missing.");

    var stops = gsLst.Elements<DocumentFormat.OpenXml.Drawing.GradientStop>().ToList();
    if (stops.Count < 2) throw new Exception($"Expected at least 2 stops, got {stops.Count}.");

    if (stops[0].Position?.Value != 0) throw new Exception($"Expected stop 0 pos=0, got {stops[0].Position?.Value}.");
    if (stops[1].Position?.Value != 100000) throw new Exception($"Expected stop 1 pos=100000, got {stops[1].Position?.Value}.");

    var lin = gradFill.Elements<DocumentFormat.OpenXml.Drawing.LinearGradientFill>().FirstOrDefault()
        ?? throw new Exception("LinearGradientFill is missing.");
    if (lin.Angle?.Value != 5400000) throw new Exception($"Expected angle=5400000, got {lin.Angle?.Value}.");

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
