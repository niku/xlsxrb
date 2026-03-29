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

    // Check table part exists
    var tableParts = wsPart.TableDefinitionParts.ToList();
    if (tableParts.Count < 1)
        throw new Exception("SCENARIO_FAIL: no table parts found");

    var table = tableParts[0].Table;
    if (table.TotalsRowCount == null || table.TotalsRowCount.Value != 1U)
        throw new Exception("SCENARIO_FAIL: totalsRowCount expected 1, got " + table.TotalsRowCount);

    var cols = table.TableColumns.Elements<TableColumn>().ToList();
    if (cols.Count < 3)
        throw new Exception("SCENARIO_FAIL: expected at least 3 columns, got " + cols.Count);

    // Check totalsRowFunction on Price column
    if (cols[1].TotalsRowFunction == null || cols[1].TotalsRowFunction.Value != TotalsRowFunctionValues.Sum)
        throw new Exception("SCENARIO_FAIL: col[1] totalsRowFunction expected sum");

    // Check calculatedColumnFormula on Tax column
    var calcFormula = cols[2].Elements<CalculatedColumnFormula>().FirstOrDefault();
    if (calcFormula == null)
        throw new Exception("SCENARIO_FAIL: col[2] missing calculatedColumnFormula");
    if (calcFormula.Text != "[Price]*0.1")
        throw new Exception("SCENARIO_FAIL: calc formula expected '[Price]*0.1', got '" + calcFormula.Text + "'");

    // Check tableStyleInfo
    var styleInfo = table.TableStyleInfo;
    if (styleInfo == null)
        throw new Exception("SCENARIO_FAIL: missing tableStyleInfo");
    if (styleInfo.Name != "TableStyleLight1")
        throw new Exception("SCENARIO_FAIL: style name expected 'TableStyleLight1', got '" + styleInfo.Name + "'");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
