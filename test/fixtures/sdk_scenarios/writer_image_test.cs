// Validates that an XLSX contains an image in the first sheet's drawing.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    var imageParts = drawingsPart.ImageParts.ToList();
    if (imageParts.Count == 0)
        throw new Exception("No image parts found in drawing.");

    // Verify at least one image part has content.
    foreach (var ip in imageParts)
    {
        using var stream = ip.GetStream();
        if (stream.Length == 0)
            throw new Exception("Image part is empty.");
    }
}
finally
{
    document.Dispose();
}
