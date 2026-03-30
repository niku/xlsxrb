var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var ws = worksheetPart.Worksheet;
    var dc = ws.GetFirstChild<DataConsolidate>();
    if (dc == null) throw new Exception("DataConsolidate element not found.");
    if (dc.Function == null || dc.Function.Value != DataConsolidateFunctionValues.Average)
        throw new Exception($"Expected function=average but got {dc.Function?.Value}.");
    if (dc.StartLabels == null || dc.StartLabels.Value != true)
        throw new Exception($"Expected startLabels=true but got {dc.StartLabels?.Value}.");
    if (dc.Link == null || dc.Link.Value != true)
        throw new Exception($"Expected link=true but got {dc.Link?.Value}.");
    var dataRefs = dc.GetFirstChild<DataReferences>();
    if (dataRefs == null) throw new Exception("DataRefs element not found.");
    var refs = dataRefs.Elements<DataReference>().ToList();
    if (refs.Count != 2) throw new Exception($"Expected 2 data refs but got {refs.Count}.");
    if (refs[0].Reference == null || refs[0].Reference.Value != "A1:B10")
        throw new Exception($"Expected ref=A1:B10 but got {refs[0].Reference?.Value}.");
    if (refs[0].Sheet == null || refs[0].Sheet.Value != "Sheet1")
        throw new Exception($"Expected sheet=Sheet1 but got {refs[0].Sheet?.Value}.");
}
finally
{
    document.Dispose();
}
