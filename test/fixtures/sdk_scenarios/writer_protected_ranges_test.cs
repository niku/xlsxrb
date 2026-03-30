var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var ws = worksheetPart.Worksheet;
    var protectedRanges = ws.Elements<ProtectedRanges>().FirstOrDefault();
    if (protectedRanges == null) throw new Exception("ProtectedRanges element not found.");
    var ranges = protectedRanges.Elements<ProtectedRange>().ToList();
    if (ranges.Count != 2) throw new Exception($"Expected 2 protected ranges but got {ranges.Count}.");
    if (ranges[0].Name == null || ranges[0].Name.Value != "EditArea")
        throw new Exception($"Expected name=EditArea but got {ranges[0].Name?.Value}.");
    if (ranges[0].SequenceOfReferences == null || !ranges[0].SequenceOfReferences.Items.Contains("A1:B10"))
        throw new Exception("Expected sqref contains A1:B10.");
    if (ranges[1].AlgorithmName == null || ranges[1].AlgorithmName.Value != "SHA-512")
        throw new Exception($"Expected algorithmName=SHA-512 but got {ranges[1].AlgorithmName?.Value}.");
}
finally
{
    document.Dispose();
}
