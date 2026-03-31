// Validates that an XLSX shape has a:rPr with font attributes in the text body.
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

    var run = para.Elements<DocumentFormat.OpenXml.Drawing.Run>().FirstOrDefault()
        ?? throw new Exception("Run is missing.");

    var rPr = run.Elements<DocumentFormat.OpenXml.Drawing.RunProperties>().FirstOrDefault()
        ?? throw new Exception("RunProperties is missing.");

    if (rPr.Bold?.Value != true) throw new Exception("Expected Bold=true.");
    if (rPr.Italic?.Value != true) throw new Exception("Expected Italic=true.");
    if (rPr.FontSize?.Value != 1400) throw new Exception($"Expected FontSize=1400, got {rPr.FontSize?.Value}.");

    var solidFill = rPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault()
        ?? throw new Exception("SolidFill in rPr is missing.");
    var srgbClr = solidFill.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()
        ?? throw new Exception("RgbColorModelHex is missing.");
    if (srgbClr.Val?.Value != "FF0000") throw new Exception($"Expected color FF0000, got {srgbClr.Val?.Value}.");

    var latin = rPr.Elements<DocumentFormat.OpenXml.Drawing.LatinFont>().FirstOrDefault()
        ?? throw new Exception("LatinFont is missing.");
    if (latin.Typeface?.Value != "Arial") throw new Exception($"Expected typeface Arial, got {latin.Typeface?.Value}.");

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
