// Generates an XLSX with an image that has alpha transparency via alphaModFix.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData(
    new Row(new Cell { CellReference = "A1", CellValue = new CellValue("test") })
));

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

// Create a minimal PNG image part
var imagePart = drawingsPart.AddNewPart<ImagePart>("image/png", "rId1");
var pngBytes = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
using (var stream = imagePart.GetStream()) { stream.Write(pngBytes, 0, pngBytes.Length); }

var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("5"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("10"),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")));

var pic = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
pic.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2u, Name = "Pic 1" },
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties(
        new DocumentFormat.OpenXml.Drawing.PictureLocks { NoChangeAspect = true })
));

// BlipFill with alphaModFix
var blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();
var blip = new DocumentFormat.OpenXml.Drawing.Blip { Embed = "rId1" };
blip.Append(new DocumentFormat.OpenXml.Drawing.AlphaModulationFixed { Amount = 75000 });
blipFill.Append(blip);
blipFill.Append(new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle()));
pic.Append(blipFill);

pic.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
        new DocumentFormat.OpenXml.Drawing.AdjustValueList()) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }));

anchor.Append(pic);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
drawingsPart.WorksheetDrawing = wsDr;

worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

var sheets = new Sheets();
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1u, Name = "Sheet1" });
workbookPart.Workbook.Append(sheets);

doc.Dispose();
