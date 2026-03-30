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
    var autoFilter = worksheetPart.Worksheet.GetFirstChild<AutoFilter>()
        ?? throw new Exception("AutoFilter is missing.");

    var filterCols = autoFilter.Elements<FilterColumn>().ToList();
    if (filterCols.Count < 2)
        throw new Exception($"Expected at least 2 filterColumns but got {filterCols.Count}.");

    // Check colorFilter
    var colorFilter = filterCols[0].GetFirstChild<ColorFilter>();
    if (colorFilter == null)
        throw new Exception("ColorFilter not found in first filterColumn.");
    if (colorFilter.FormatId?.Value != 0U)
        throw new Exception($"Expected colorFilter dxfId=0 but got {colorFilter.FormatId?.Value}.");

    // Check iconFilter
    var iconFilter = filterCols[1].GetFirstChild<IconFilter>();
    if (iconFilter == null)
        throw new Exception("IconFilter not found in second filterColumn.");
    if (iconFilter.IconSet?.Value != IconSetValues.ThreeArrows)
        throw new Exception($"Expected iconSet=3Arrows but got {iconFilter.IconSet?.Value}.");
    if (iconFilter.IconId?.Value != 1U)
        throw new Exception($"Expected iconId=1 but got {iconFilter.IconId?.Value}.");
}
finally
{
    document.Dispose();
}
