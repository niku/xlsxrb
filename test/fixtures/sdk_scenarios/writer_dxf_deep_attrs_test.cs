var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var stylesPart = workbookPart.WorkbookStylesPart;
var stylesheet = stylesPart.Stylesheet;

var dxfs = stylesheet.DifferentialFormats;
if (dxfs == null)
    throw new Exception("SCENARIO_FAIL: No dxfs element found");

var dxfList = dxfs.Elements<DifferentialFormat>().ToList();
if (dxfList.Count == 0)
    throw new Exception("SCENARIO_FAIL: No dxf elements found");

var dxf = dxfList[0];

// Check font
var font = dxf.Font;
if (font == null)
    throw new Exception("SCENARIO_FAIL: DXF font is null");

// Check numFmt
var numFmt = dxf.NumberingFormat;
if (numFmt == null)
    throw new Exception("SCENARIO_FAIL: DXF numFmt is null");
if (numFmt.FormatCode?.Value != "#,##0.00")
    throw new Exception($"SCENARIO_FAIL: numFmt formatCode expected '#,##0.00', got '{numFmt.FormatCode?.Value}'");

// Check alignment
var alignment = dxf.Alignment;
if (alignment == null)
    throw new Exception("SCENARIO_FAIL: DXF alignment is null");
if (alignment.Horizontal?.Value != HorizontalAlignmentValues.Center)
    throw new Exception($"SCENARIO_FAIL: alignment horizontal expected Center, got {alignment.Horizontal?.Value}");
if (alignment.WrapText?.Value != true)
    throw new Exception($"SCENARIO_FAIL: alignment wrapText expected true, got {alignment.WrapText?.Value}");

// Check protection
var protection = dxf.Protection;
if (protection == null)
    throw new Exception("SCENARIO_FAIL: DXF protection is null");
if (protection.Locked?.Value != false)
    throw new Exception($"SCENARIO_FAIL: protection locked expected false, got {protection.Locked?.Value}");
if (protection.Hidden?.Value != true)
    throw new Exception($"SCENARIO_FAIL: protection hidden expected true, got {protection.Hidden?.Value}");

doc.Dispose();
