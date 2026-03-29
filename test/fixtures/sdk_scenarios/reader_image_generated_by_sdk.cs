// Creates an XLSX with an embedded image (1x1 white pixel PNG).
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("data")) })
);
worksheetPart.Worksheet = new Worksheet(sheetData);

// Create drawing part.
var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

// Minimal 1x1 white pixel PNG.
byte[] pngBytes = new byte[] {
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
    0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
    0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
    0x54, 0x08, 0xD7, 0x63, 0xF8, 0xCF, 0xC0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xE2, 0x21, 0xBC,
    0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E,
    0x44, 0xAE, 0x42, 0x60, 0x82
};

var imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
using (var ms = new MemoryStream(pngBytes))
{
    imagePart.FeedData(ms);
}

string imageRelId = drawingsPart.GetIdOfPart(imagePart);

// Build drawing XML with twoCellAnchor + pic.
var wsDr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();
var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor();

var fromMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker();
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"));
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"));
fromMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));

var toMarker = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker();
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("5"));
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"));
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("10"));
toMarker.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0"));

anchor.Append(fromMarker);
anchor.Append(toMarker);

var pic = new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture();
var nvPicPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties();
nvPicPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties { Id = 2, Name = "TestImage" });
nvPicPr.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties());
pic.Append(nvPicPr);

var blipFill = new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill();
blipFill.Append(new DocumentFormat.OpenXml.Drawing.Blip { Embed = imageRelId });
blipFill.Append(new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle()));
pic.Append(blipFill);

var spPr = new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties();
spPr.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle });
pic.Append(spPr);

anchor.Append(pic);
anchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData());
wsDr.Append(anchor);
drawingsPart.WorksheetDrawing = wsDr;
drawingsPart.WorksheetDrawing.Save();

// Add <drawing> element to worksheet.
var drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
worksheetPart.Worksheet.Append(new Drawing { Id = drawingRelId });
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
