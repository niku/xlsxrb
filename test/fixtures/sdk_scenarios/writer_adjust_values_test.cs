// Validates that an XLSX shape has a:avLst with a:gd adjust values.
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

    var prstGeom = spPr.Elements<DocumentFormat.OpenXml.Drawing.PresetGeometry>().FirstOrDefault()
        ?? throw new Exception("PresetGeometry is missing.");

    if (prstGeom.Preset?.Value != DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle)
        throw new Exception($"Expected roundRect but got {prstGeom.Preset?.Value}.");

    var avLst = prstGeom.Elements<DocumentFormat.OpenXml.Drawing.AdjustValueList>().FirstOrDefault()
        ?? throw new Exception("AdjustValueList is missing.");

    var gds = avLst.Elements<DocumentFormat.OpenXml.Drawing.ShapeGuide>().ToList();
    if (gds.Count == 0) throw new Exception("No ShapeGuide (gd) elements found.");

    if (gds[0].Name?.Value != "adj")
        throw new Exception($"Expected gd name 'adj' but got '{gds[0].Name?.Value}'.");
    if (gds[0].Formula?.Value != "val 16667")
        throw new Exception($"Expected gd fmla 'val 16667' but got '{gds[0].Formula?.Value}'.");

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
