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
    var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() ?? throw new Exception("SheetData is missing.");
    var row = sheetData.Elements<Row>().First();
    var cell = row.Elements<Cell>().First();

    // Cell should have a numeric value (fractional serial)
    if (cell.DataType != null && cell.DataType.Value != CellValues.Number)
        throw new Exception($"Expected numeric cell but got DataType={cell.DataType?.Value}.");

    var rawValue = double.Parse(cell.CellValue?.Text ?? "0");
    // 2024-03-15 14:30:00 = serial 45366.604166...
    if (Math.Abs(rawValue - 45366.604166666664) > 0.001)
        throw new Exception($"Expected serial ~45366.6042 but got {rawValue}.");

    // Verify the cell has a style index pointing to a datetime format
    if (cell.StyleIndex == null)
        throw new Exception("Cell has no StyleIndex.");

    var stylesPart = workbookPart.WorkbookStylesPart ?? throw new Exception("WorkbookStylesPart is missing.");
    var styleSheet = stylesPart.Stylesheet ?? throw new Exception("Stylesheet is missing.");

    var xfIndex = (int)cell.StyleIndex.Value;
    var cellXfs = styleSheet.CellFormats?.Elements<CellFormat>().ToList()
        ?? throw new Exception("CellFormats are missing.");
    if (xfIndex >= cellXfs.Count)
        throw new Exception($"StyleIndex {xfIndex} out of range (total xfs: {cellXfs.Count}).");

    var xf = cellXfs[xfIndex];
    var numFmtId = xf.NumberFormatId?.Value ?? 0;
    if (numFmtId < 164)
        throw new Exception($"Expected custom numFmtId >= 164 but got {numFmtId}.");

    // Verify the numFmt contains datetime-like pattern
    var numFmts = styleSheet.NumberingFormats?.Elements<NumberingFormat>().ToList();
    if (numFmts == null || numFmts.Count == 0)
        throw new Exception("No custom number formats found.");

    var fmt = numFmts.FirstOrDefault(nf => nf.NumberFormatId?.Value == numFmtId);
    if (fmt == null)
        throw new Exception($"NumFmt with id {numFmtId} not found.");
    if (!fmt.FormatCode?.Value?.Contains("hh:mm:ss") ?? true)
        throw new Exception($"Expected datetime format code with hh:mm:ss but got {fmt.FormatCode?.Value}.");
}
finally
{
    document.Dispose();
}
