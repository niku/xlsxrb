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
    var fills = styles.Fills;

    // Fill[2] should be gradient (0=none, 1=gray125, 2=gradient)
    var fill = fills.Elements<Fill>().ElementAt(2);
    var gf = fill.GetFirstChild<GradientFill>();
    if (gf == null)
        throw new Exception("SCENARIO_FAIL: gradientFill not found in fill[2]");

    if (gf.Type == null || gf.Type.Value != GradientValues.Linear)
        throw new Exception("SCENARIO_FAIL: expected type=linear, got " + gf.Type);

    if (gf.Degree == null || gf.Degree.Value != 90.0)
        throw new Exception("SCENARIO_FAIL: expected degree=90, got " + gf.Degree);

    var stops = gf.Elements<GradientStop>().ToList();
    if (stops.Count != 2)
        throw new Exception("SCENARIO_FAIL: expected 2 stops, got " + stops.Count);

    if (stops[0].Position == null || stops[0].Position.Value != 0.0)
        throw new Exception("SCENARIO_FAIL: expected stop[0] position=0, got " + stops[0].Position);

    var color0 = stops[0].GetFirstChild<Color>();
    if (color0 == null || color0.Rgb == null || color0.Rgb.Value != "FFFF0000")
        throw new Exception("SCENARIO_FAIL: expected stop[0] color=FFFF0000, got " + color0?.Rgb);

    if (stops[1].Position == null || stops[1].Position.Value != 1.0)
        throw new Exception("SCENARIO_FAIL: expected stop[1] position=1, got " + stops[1].Position);

    var color1 = stops[1].GetFirstChild<Color>();
    if (color1 == null || color1.Rgb == null || color1.Rgb.Value != "FF0000FF")
        throw new Exception("SCENARIO_FAIL: expected stop[1] color=FF0000FF, got " + color1?.Rgb);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
