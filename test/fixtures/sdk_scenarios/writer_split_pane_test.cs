var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var validator = new OpenXmlValidator(FileFormatVersions.Office2007);
    var validationErrors = validator.Validate(document).Take(10).ToList();
    if (validationErrors.Any())
    {
        var message = string.Join(Environment.NewLine, validationErrors.Select(e => e.Description));
        throw new Exception($"OpenXmlValidator reported errors:{Environment.NewLine}{message}");
    }

    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var worksheetPart = workbookPart.WorksheetParts.First();
    var sheetViews = worksheetPart.Worksheet.GetFirstChild<SheetViews>()
        ?? throw new Exception("SheetViews is missing.");
    var sheetView = sheetViews.GetFirstChild<SheetView>()
        ?? throw new Exception("SheetView is missing.");
    var pane = sheetView.GetFirstChild<Pane>()
        ?? throw new Exception("Pane is missing.");

    if (pane.HorizontalSplit?.Value != 2400D)
        throw new Exception($"Expected xSplit=2400 but got {pane.HorizontalSplit?.Value}.");
    if (pane.VerticalSplit?.Value != 1800D)
        throw new Exception($"Expected ySplit=1800 but got {pane.VerticalSplit?.Value}.");
    if (pane.TopLeftCell?.Value != "C4")
        throw new Exception($"Expected topLeftCell=C4 but got {pane.TopLeftCell?.Value}.");
    if (pane.ActivePane?.Value != PaneValues.BottomRight)
        throw new Exception($"Expected activePane=bottomRight but got {pane.ActivePane?.Value}.");
    if (pane.State != null && pane.State.Value == PaneStateValues.Frozen)
        throw new Exception("Expected split pane (not frozen) but state is frozen.");
}
finally
{
    document.Dispose();
}
