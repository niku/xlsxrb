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

    // Font[1] should have theme color
    var font = fonts.Elements<Font>().ElementAt(1);
    var color = font.GetFirstChild<Color>();
    if (color == null)
        throw new Exception("SCENARIO_FAIL: color element not found");

    if (color.Theme == null || color.Theme.Value != 1U)
        throw new Exception("SCENARIO_FAIL: expected theme=1, got " + color.Theme);

    if (color.Tint == null || Math.Abs(color.Tint.Value - (-0.25)) > 0.001)
        throw new Exception("SCENARIO_FAIL: expected tint=-0.25, got " + color.Tint);

    // Font[2] should have indexed color
    var font2 = fonts.Elements<Font>().ElementAt(2);
    var color2 = font2.GetFirstChild<Color>();
    if (color2 == null)
        throw new Exception("SCENARIO_FAIL: color2 element not found");

    if (color2.Indexed == null || color2.Indexed.Value != 10U)
        throw new Exception("SCENARIO_FAIL: expected indexed=10, got " + color2.Indexed);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
