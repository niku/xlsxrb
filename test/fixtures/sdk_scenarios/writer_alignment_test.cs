using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(10).ToList();
    if (errors.Count > 0)
    {
        foreach (var e in errors)
            Console.Error.WriteLine("Validation: " + e.Description);
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count);
    }

    var wbPart = document.WorkbookPart;
    var styles = wbPart.WorkbookStylesPart.Stylesheet;
    var cellXfs = styles.CellFormats;

    // Find the xf with alignment (index 1)
    var xf = cellXfs.Elements<CellFormat>().ElementAt(1);
    if (xf.ApplyAlignment == null || !xf.ApplyAlignment.Value)
        throw new Exception("SCENARIO_FAIL: applyAlignment not set on xf[1]");

    var alignment = xf.Alignment;
    if (alignment == null)
        throw new Exception("SCENARIO_FAIL: alignment element missing on xf[1]");

    if (alignment.Horizontal == null || alignment.Horizontal.Value != HorizontalAlignmentValues.Center)
        throw new Exception("SCENARIO_FAIL: expected horizontal=center, got " + alignment.Horizontal);

    if (alignment.Vertical == null || alignment.Vertical.Value != VerticalAlignmentValues.Top)
        throw new Exception("SCENARIO_FAIL: expected vertical=top, got " + alignment.Vertical);

    if (alignment.WrapText == null || !alignment.WrapText.Value)
        throw new Exception("SCENARIO_FAIL: expected wrapText=true");

    if (alignment.TextRotation == null || alignment.TextRotation.Value != 45U)
        throw new Exception("SCENARIO_FAIL: expected textRotation=45, got " + alignment.TextRotation);

    if (alignment.Indent == null || alignment.Indent.Value != 2U)
        throw new Exception("SCENARIO_FAIL: expected indent=2, got " + alignment.Indent);

    if (alignment.ShrinkToFit == null || !alignment.ShrinkToFit.Value)
        throw new Exception("SCENARIO_FAIL: expected shrinkToFit=true");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
