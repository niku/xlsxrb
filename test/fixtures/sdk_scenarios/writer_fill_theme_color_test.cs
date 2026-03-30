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

    // Fill[2] should have theme fg color
    var fill = fills.Elements<Fill>().ElementAt(2);
    var pf = fill.GetFirstChild<PatternFill>();
    if (pf == null)
        throw new Exception("SCENARIO_FAIL: patternFill not found");

    var fg = pf.GetFirstChild<ForegroundColor>();
    if (fg == null)
        throw new Exception("SCENARIO_FAIL: fgColor not found");

    if (fg.Theme == null || fg.Theme.Value != 4U)
        throw new Exception("SCENARIO_FAIL: expected theme=4, got " + fg.Theme);

    if (fg.Tint == null || Math.Abs(fg.Tint.Value - 0.6) > 0.001)
        throw new Exception("SCENARIO_FAIL: expected tint=0.6, got " + fg.Tint);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
