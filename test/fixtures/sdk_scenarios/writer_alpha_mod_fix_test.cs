// Validates that an XLSX image has a:alphaModFix on the blip for transparency.
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

    var blip = blipFill.Elements<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault()
        ?? throw new Exception("Blip is missing.");

    var alphaMod = blip.Elements<DocumentFormat.OpenXml.Drawing.AlphaModulationFixed>().FirstOrDefault()
        ?? throw new Exception("AlphaModulationFixed is missing.");

    if (alphaMod.Amount?.Value != 50000) throw new Exception($"Expected amt=50000, got {alphaMod.Amount?.Value}.");

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
