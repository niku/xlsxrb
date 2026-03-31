// Validates that a table column with totalsRowFormula is correctly written.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);
    var tableParts = wsPart.TableDefinitionParts.ToList();
    if (tableParts.Count < 1)
        throw new Exception("SCENARIO_FAIL: no table parts found");

    var cols = tableParts[0].Table.TableColumns.Elements<TableColumn>().ToList();
    if (cols.Count < 2)
        throw new Exception("SCENARIO_FAIL: expected at least 2 columns, got " + cols.Count);

    var formula = cols[1].Elements<TotalsRowFormula>().FirstOrDefault();
    if (formula == null)
        throw new Exception("SCENARIO_FAIL: TotalsRowFormula missing on column 1");
    if (formula.Text != "SUBTOTAL(109,[Price])")
        throw new Exception("SCENARIO_FAIL: expected 'SUBTOTAL(109,[Price])', got '" + formula.Text + "'");

    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).Take(5).ToList();
    if (errors.Count > 0)
        throw new Exception("SCENARIO_FAIL: validation errors: " + errors.Count);

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
