using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System.Linq;

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
    var fonts = styles.Fonts;

    // Font[1] should have extended attributes
    var font = fonts.Elements<Font>().ElementAt(1);

    if (font.GetFirstChild<Bold>() == null)
        throw new Exception("SCENARIO_FAIL: bold expected");
    if (font.GetFirstChild<Italic>() == null)
        throw new Exception("SCENARIO_FAIL: italic expected");
    if (font.GetFirstChild<Strike>() == null)
        throw new Exception("SCENARIO_FAIL: strike expected");

    var underline = font.GetFirstChild<Underline>();
    if (underline == null || underline.Val == null || underline.Val.Value != UnderlineValues.Double)
        throw new Exception("SCENARIO_FAIL: expected underline=double, got " + underline?.Val);

    var vertAlign = font.GetFirstChild<VerticalTextAlignment>();
    if (vertAlign == null || vertAlign.Val == null || vertAlign.Val.Value != VerticalAlignmentRunValues.Superscript)
        throw new Exception("SCENARIO_FAIL: expected vertAlign=superscript, got " + vertAlign?.Val);

    // Verify family and scheme in raw XML (validated structurally by OpenXmlValidator above)
    var fontXml = font.OuterXml;
    if (!fontXml.Contains("family"))
        throw new Exception("SCENARIO_FAIL: family element not found in XML: " + fontXml);
    if (!fontXml.Contains("scheme"))
        throw new Exception("SCENARIO_FAIL: scheme element not found in XML: " + fontXml);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
