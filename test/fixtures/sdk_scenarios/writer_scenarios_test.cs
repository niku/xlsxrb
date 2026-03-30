var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var ws = worksheetPart.Worksheet;
    var scenarios = ws.GetFirstChild<Scenarios>();
    if (scenarios == null) throw new Exception("Scenarios element not found.");
    if (scenarios.Current == null || scenarios.Current.Value != 0U)
        throw new Exception($"Expected current=0 but got {scenarios.Current?.Value}.");
    if (scenarios.Show == null || scenarios.Show.Value != 0U)
        throw new Exception($"Expected show=0 but got {scenarios.Show?.Value}.");
    var scenarioList = scenarios.Elements<Scenario>().ToList();
    if (scenarioList.Count != 1) throw new Exception($"Expected 1 scenario but got {scenarioList.Count}.");
    var sc = scenarioList[0];
    if (sc.Name == null || sc.Name.Value != "Best Case")
        throw new Exception($"Expected name=Best Case but got {sc.Name?.Value}.");
    if (sc.User == null || sc.User.Value != "Admin")
        throw new Exception($"Expected user=Admin but got {sc.User?.Value}.");
    var inputs = sc.Elements<InputCells>().ToList();
    if (inputs.Count != 2) throw new Exception($"Expected 2 input cells but got {inputs.Count}.");
    if (inputs[0].CellReference == null || inputs[0].CellReference.Value != "A1")
        throw new Exception($"Expected r=A1 but got {inputs[0].CellReference?.Value}.");
    if (inputs[0].Val == null || inputs[0].Val.Value != "200")
        throw new Exception($"Expected val=200 but got {inputs[0].Val?.Value}.");
}
finally
{
    document.Dispose();
}
