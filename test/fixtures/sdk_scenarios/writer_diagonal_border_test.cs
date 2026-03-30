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
    var borders = styles.Borders;

    var border = borders.Elements<Border>().ElementAt(1);
    if (border.DiagonalUp == null || !border.DiagonalUp.Value)
        throw new Exception("SCENARIO_FAIL: diagonalUp not set");
    if (border.DiagonalDown == null || !border.DiagonalDown.Value)
        throw new Exception("SCENARIO_FAIL: diagonalDown not set");

    var diag = border.GetFirstChild<DiagonalBorder>();
    if (diag == null || diag.Style == null || diag.Style.Value != BorderStyleValues.Thin)
        throw new Exception("SCENARIO_FAIL: expected diagonal style=thin, got " + diag?.Style);

    var color = diag.GetFirstChild<Color>();
    if (color == null || color.Rgb == null || color.Rgb.Value != "FFFF0000")
        throw new Exception("SCENARIO_FAIL: expected diagonal color=FFFF0000, got " + color?.Rgb);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
