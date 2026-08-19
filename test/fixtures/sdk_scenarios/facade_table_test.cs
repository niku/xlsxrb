using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var errors = validator.Validate(document).ToList();
    if (errors.Count > 0)
    {
        var messages = string.Join("\n", errors.Select(e => $"  - {e.Description} (Part: {e.Part?.Uri})"));
        throw new Exception($"OpenXML Validation errors:\n{messages}");
    }

    var wbPart = document.WorkbookPart ?? throw new Exception("WorkbookPart missing");
    var sheet = wbPart.Workbook.Sheets.Elements<Sheet>().First();
    var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id);

    var tableParts = wsPart.TableDefinitionParts.ToList();
    if (tableParts.Count < 1)
        throw new Exception("SCENARIO_FAIL: TableDefinitionParts missing");

    var table = tableParts[0].Table;
    if (table.DisplayName != "EmployeesTable" && table.Name != "EmployeesTable")
        throw new Exception($"SCENARIO_FAIL: Table name mismatch, got {table.Name}");

    var cols = table.TableColumns.Elements<TableColumn>().ToList();
    if (cols.Count != 4)
        throw new Exception($"SCENARIO_FAIL: Expected 4 columns, got {cols.Count}");

    if (table.TableStyleInfo?.Name != "TableStyleMedium9")
        throw new Exception($"SCENARIO_FAIL: Table style expected TableStyleMedium9, got {table.TableStyleInfo?.Name}");

    Console.Error.WriteLine("SCENARIO_PASS");
}
finally
{
    document.Dispose();
}
