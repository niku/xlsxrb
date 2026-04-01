// Validates that an XLSX chart title has formatted text with a:rPr attributes.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var drawingsPart = worksheetPart.DrawingsPart
        ?? throw new Exception("DrawingsPart is missing.");

    var chartPart = drawingsPart.ChartParts.FirstOrDefault()
        ?? throw new Exception("ChartPart is missing.");

    var chart = chartPart.ChartSpace?.Elements<DocumentFormat.OpenXml.Drawing.Charts.Chart>().FirstOrDefault()
        ?? throw new Exception("Chart is missing.");

    var title = chart.Title ?? throw new Exception("Title is missing.");
    var tx = title.ChartText ?? throw new Exception("ChartText is missing.");
    var rich = tx.RichText ?? throw new Exception("RichText is missing.");
    var para = rich.Elements<DocumentFormat.OpenXml.Drawing.Paragraph>().FirstOrDefault()
        ?? throw new Exception("Paragraph is missing.");
    var run = para.Elements<DocumentFormat.OpenXml.Drawing.Run>().FirstOrDefault()
        ?? throw new Exception("Run is missing.");
    var rPr = run.RunProperties ?? throw new Exception("RunProperties is missing.");

    if (rPr.Bold?.Value != true) throw new Exception($"Expected bold=true, got {rPr.Bold?.Value}.");
    if (rPr.Italic?.Value != true) throw new Exception($"Expected italic=true, got {rPr.Italic?.Value}.");
    if (rPr.FontSize?.Value != 1400) throw new Exception($"Expected sz=1400, got {rPr.FontSize?.Value}.");

    var solidFill = rPr.Elements<DocumentFormat.OpenXml.Drawing.SolidFill>().FirstOrDefault()
        ?? throw new Exception("SolidFill is missing.");
    var srgbClr = solidFill.Elements<DocumentFormat.OpenXml.Drawing.RgbColorModelHex>().FirstOrDefault()
        ?? throw new Exception("RgbColorModelHex is missing.");
    if (srgbClr.Val?.Value != "FF0000") throw new Exception($"Expected color FF0000, got {srgbClr.Val?.Value}.");

    var latin = rPr.Elements<DocumentFormat.OpenXml.Drawing.LatinFont>().FirstOrDefault()
        ?? throw new Exception("LatinFont is missing.");
    if (latin.Typeface?.Value != "Arial") throw new Exception($"Expected typeface Arial, got {latin.Typeface?.Value}.");

    var text = run.Text?.Text;
    if (text != "My Chart") throw new Exception($"Expected text 'My Chart', got '{text}'.");

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
