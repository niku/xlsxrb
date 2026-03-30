var doc = SpreadsheetDocument.Open(XlsxPath, false);

var workbookPart = doc.WorkbookPart;
var ssp = workbookPart.SharedStringTablePart;
if (ssp == null)
    throw new Exception("SCENARIO_FAIL: No SharedStringTablePart found");

var sst = ssp.SharedStringTable;
var items = sst.Elements<SharedStringItem>().ToList();
if (items.Count == 0)
    throw new Exception("SCENARIO_FAIL: No shared string items found");

// Check the first shared string item (rich text with extended attributes)
var si = items[0];
var runs = si.Elements<Run>().ToList();
if (runs.Count != 4)
    throw new Exception($"SCENARIO_FAIL: Expected 4 runs, got {runs.Count}");

// Run 0: strike
var rpr0 = runs[0].RunProperties;
if (rpr0 == null || rpr0.GetFirstChild<Strike>() == null)
    throw new Exception("SCENARIO_FAIL: Run0 expected <strike/>");

// Run 1: underline double
var rpr1 = runs[1].RunProperties;
if (rpr1 == null)
    throw new Exception("SCENARIO_FAIL: Run1 RunProperties is null");
var u1 = rpr1.GetFirstChild<Underline>();
if (u1 == null || u1.Val?.Value != UnderlineValues.Double)
    throw new Exception($"SCENARIO_FAIL: Run1 expected underline double, got {u1?.Val?.Value}");

// Run 2: vertAlign superscript
var rpr2 = runs[2].RunProperties;
if (rpr2 == null)
    throw new Exception("SCENARIO_FAIL: Run2 RunProperties is null");
var va2 = rpr2.GetFirstChild<VerticalTextAlignment>();
if (va2 == null || va2.Val?.Value != VerticalAlignmentRunValues.Superscript)
    throw new Exception($"SCENARIO_FAIL: Run2 expected superscript, got {va2?.Val?.Value}");

// Run 3: theme color, family, scheme
var rpr3 = runs[3].RunProperties;
if (rpr3 == null)
    throw new Exception("SCENARIO_FAIL: Run3 RunProperties is null");
var color3 = rpr3.GetFirstChild<Color>();
if (color3 == null || color3.Theme?.Value != 1U)
    throw new Exception($"SCENARIO_FAIL: Run3 expected theme=1, got {color3?.Theme?.Value}");

// Check family and scheme via OuterXml since SDK strongly-typed access can be unreliable
var xml3 = rpr3.OuterXml;
if (!xml3.Contains("family"))
    throw new Exception($"SCENARIO_FAIL: Run3 expected family in XML: {xml3}");
if (!xml3.Contains("scheme"))
    throw new Exception($"SCENARIO_FAIL: Run3 expected scheme in XML: {xml3}");

doc.Dispose();
