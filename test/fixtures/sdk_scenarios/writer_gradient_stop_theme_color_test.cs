var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var stylesPart = workbookPart.WorkbookStylesPart;
var stylesheet = stylesPart.Stylesheet;

var fills = stylesheet.Fills.Elements<Fill>().ToList();
// Find the gradient fill (skip default fills at index 0 and 1)
Fill gradFill = null;
foreach (var fill in fills)
{
    if (fill.GradientFill != null)
    {
        gradFill = fill;
        break;
    }
}
if (gradFill == null)
    throw new Exception("SCENARIO_FAIL: No gradient fill found");

var gf = gradFill.GradientFill;
var stops = gf.Elements<GradientStop>().ToList();
if (stops.Count != 2)
    throw new Exception($"SCENARIO_FAIL: Expected 2 stops, got {stops.Count}");

// Stop 0: theme color
var c0 = stops[0].Color;
if (c0 == null || c0.Theme?.Value != 4U)
    throw new Exception($"SCENARIO_FAIL: Stop0 theme expected 4, got {c0?.Theme?.Value}");
if (c0.Tint?.Value == null || Math.Abs(c0.Tint.Value - (-0.5)) > 0.01)
    throw new Exception($"SCENARIO_FAIL: Stop0 tint expected -0.5, got {c0.Tint?.Value}");

// Stop 1: indexed color
var c1 = stops[1].Color;
if (c1 == null || c1.Indexed?.Value != 12U)
    throw new Exception($"SCENARIO_FAIL: Stop1 indexed expected 12, got {c1?.Indexed?.Value}");

doc.Dispose();
