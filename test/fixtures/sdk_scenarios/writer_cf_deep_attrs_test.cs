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
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var cfs = wsPart.Worksheet.Elements<ConditionalFormatting>().ToList();

    // We expect 2 conditional formatting blocks (dataBar + iconSet)
    if (cfs.Count < 1)
        throw new Exception("SCENARIO_FAIL: no conditional formatting found");

    // Find dataBar rule
    ConditionalFormattingRule dbRule = null;
    ConditionalFormattingRule isRule = null;
    foreach (var cf in cfs)
    {
        foreach (var rule in cf.Elements<ConditionalFormattingRule>())
        {
            if (rule.Elements<DataBar>().Any())
                dbRule = rule;
            if (rule.Elements<IconSet>().Any())
                isRule = rule;
        }
    }

    if (dbRule == null)
        throw new Exception("SCENARIO_FAIL: no dataBar rule found");

    var dataBar = dbRule.Elements<DataBar>().First();
    if (dataBar.MinLength == null || dataBar.MinLength.Value != 5U)
        throw new Exception("SCENARIO_FAIL: dataBar minLength expected 5, got " + dataBar.MinLength);
    if (dataBar.MaxLength == null || dataBar.MaxLength.Value != 90U)
        throw new Exception("SCENARIO_FAIL: dataBar maxLength expected 90, got " + dataBar.MaxLength);
    if (dataBar.ShowValue == null || dataBar.ShowValue.Value != false)
        throw new Exception("SCENARIO_FAIL: dataBar showValue expected false, got " + dataBar.ShowValue);

    if (isRule == null)
        throw new Exception("SCENARIO_FAIL: no iconSet rule found");

    var iconSet = isRule.Elements<IconSet>().First();
    if (iconSet.Reverse == null || iconSet.Reverse.Value != true)
        throw new Exception("SCENARIO_FAIL: iconSet reverse expected true, got " + iconSet.Reverse);
    if (iconSet.ShowValue == null || iconSet.ShowValue.Value != false)
        throw new Exception("SCENARIO_FAIL: iconSet showValue expected false, got " + iconSet.ShowValue);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
