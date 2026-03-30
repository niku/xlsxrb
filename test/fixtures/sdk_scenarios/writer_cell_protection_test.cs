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
    var cellXfs = styles.CellFormats;

    var xf = cellXfs.Elements<CellFormat>().ElementAt(1);
    if (xf.ApplyProtection == null || !xf.ApplyProtection.Value)
        throw new Exception("SCENARIO_FAIL: applyProtection not set");

    var prot = xf.Protection;
    if (prot == null)
        throw new Exception("SCENARIO_FAIL: protection element missing");

    if (prot.Locked == null || prot.Locked.Value != false)
        throw new Exception("SCENARIO_FAIL: expected locked=false, got " + prot.Locked);

    if (prot.Hidden == null || prot.Hidden.Value != true)
        throw new Exception("SCENARIO_FAIL: expected hidden=true, got " + prot.Hidden);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
