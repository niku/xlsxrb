var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var ws = worksheetPart.Worksheet;
    var ignoredErrors = ws.GetFirstChild<IgnoredErrors>();
    if (ignoredErrors == null) throw new Exception("IgnoredErrors element not found.");
    var errors = ignoredErrors.Elements<IgnoredError>().ToList();
    if (errors.Count != 1) throw new Exception($"Expected 1 ignored error but got {errors.Count}.");
    if (errors[0].SequenceOfReferences == null || !errors[0].SequenceOfReferences.Items.Contains("A1:B2"))
        throw new Exception("Expected sqref contains A1:B2.");
    if (errors[0].NumberStoredAsText == null || errors[0].NumberStoredAsText.Value != true)
        throw new Exception($"Expected numberStoredAsText=true but got {errors[0].NumberStoredAsText?.Value}.");
    if (errors[0].EvalError == null || errors[0].EvalError.Value != true)
        throw new Exception($"Expected evalError=true but got {errors[0].EvalError?.Value}.");
}
finally
{
    document.Dispose();
}
