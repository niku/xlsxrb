// Validates that an XLSX shape has bodyPr with autofit child elements.
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
    if (anchors.Count < 3) throw new Exception($"Expected at least 3 anchors, got {anchors.Count}.");

    // Shape 1: autofit=none -> noAutofit
    var sp1 = anchors[0].Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault()
        ?? throw new Exception("Shape1 is missing.");
    var txBody1 = sp1.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody>().FirstOrDefault()
        ?? throw new Exception("TextBody1 is missing.");
    var bodyPr1 = txBody1.Elements<DocumentFormat.OpenXml.Drawing.BodyProperties>().FirstOrDefault()
        ?? throw new Exception("BodyProperties1 is missing.");
    var noAutofit = bodyPr1.Elements<DocumentFormat.OpenXml.Drawing.NoAutoFit>().FirstOrDefault()
        ?? throw new Exception("NoAutoFit is missing on shape1.");

    // Shape 2: autofit=shape -> spAutoFit
    var sp2 = anchors[1].Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault()
        ?? throw new Exception("Shape2 is missing.");
    var txBody2 = sp2.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody>().FirstOrDefault()
        ?? throw new Exception("TextBody2 is missing.");
    var bodyPr2 = txBody2.Elements<DocumentFormat.OpenXml.Drawing.BodyProperties>().FirstOrDefault()
        ?? throw new Exception("BodyProperties2 is missing.");
    var spAutoFit = bodyPr2.Elements<DocumentFormat.OpenXml.Drawing.ShapeAutoFit>().FirstOrDefault()
        ?? throw new Exception("ShapeAutoFit is missing on shape2.");

    // Shape 3: autofit=normal -> normAutofit
    var sp3 = anchors[2].Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.Shape>().FirstOrDefault()
        ?? throw new Exception("Shape3 is missing.");
    var txBody3 = sp3.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TextBody>().FirstOrDefault()
        ?? throw new Exception("TextBody3 is missing.");
    var bodyPr3 = txBody3.Elements<DocumentFormat.OpenXml.Drawing.BodyProperties>().FirstOrDefault()
        ?? throw new Exception("BodyProperties3 is missing.");
    var normAutofit = bodyPr3.Elements<DocumentFormat.OpenXml.Drawing.NormalAutoFit>().FirstOrDefault()
        ?? throw new Exception("NormalAutoFit is missing on shape3.");

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
