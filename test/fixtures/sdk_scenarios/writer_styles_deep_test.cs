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

    // Check cellStyleXfs
    var csxfs = styles.CellStyleFormats;
    if (csxfs == null || csxfs.Count == null || csxfs.Count.Value < 2)
        throw new Exception("SCENARIO_FAIL: expected at least 2 cellStyleXfs, got " + csxfs?.Count);

    // Check cellStyles
    var cellStyles = styles.CellStyles;
    if (cellStyles == null || cellStyles.Count == null || cellStyles.Count.Value < 2)
        throw new Exception("SCENARIO_FAIL: expected at least 2 cellStyles, got " + cellStyles?.Count);

    var normalStyle = cellStyles.Elements<CellStyle>().FirstOrDefault(cs => cs.Name != null && cs.Name.Value == "Normal");
    if (normalStyle == null)
        throw new Exception("SCENARIO_FAIL: Normal style not found");

    var heading1Style = cellStyles.Elements<CellStyle>().FirstOrDefault(cs => cs.Name != null && cs.Name.Value == "Heading1");
    if (heading1Style == null)
        throw new Exception("SCENARIO_FAIL: Heading1 style not found");

    if (heading1Style.BuiltinId == null || heading1Style.BuiltinId.Value != 1U)
        throw new Exception("SCENARIO_FAIL: Heading1 builtinId expected 1, got " + heading1Style.BuiltinId);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
