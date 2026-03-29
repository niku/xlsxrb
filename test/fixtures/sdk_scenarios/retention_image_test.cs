// Validates image retention: XLSX should still contain a drawing with an image.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing after retention.");

    var imageParts = drawingsPart.ImageParts.ToList();
    if (imageParts.Count == 0)
        throw new Exception("No image parts found after retention.");

    using var stream = imageParts[0].GetStream();
    if (stream.Length == 0)
        throw new Exception("Image data lost during retention.");
}
finally
{
    document.Dispose();
}
