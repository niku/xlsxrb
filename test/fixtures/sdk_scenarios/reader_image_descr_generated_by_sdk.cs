var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = document.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData(new Row(new Cell
{
    CellReference = "A1",
    DataType = CellValues.InlineString,
    InlineString = new InlineString(new Text("Image test"))
})));

var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
drawingsPart.WorksheetDrawing = new DocumentFormat.OpenXml.Drawing.Spreadsheet.WorksheetDrawing();

var imagePart = drawingsPart.AddImagePart(ImagePartType.Png, "rId1");
var pngBytes = new byte[]
{
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
using (var stream = new MemoryStream(pngBytes))
{
    imagePart.FeedData(stream);
}

var anchor = new DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor(
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("0"),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("0"),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")
    ),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnId("5"),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ColumnOffset("0"),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowId("10"),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.RowOffset("0")
    ),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture(
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureProperties(
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties
            {
                Id = 1U,
                Name = "DescribedImage",
                Description = "Image alt text from SDK"
            },
            new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualPictureDrawingProperties(
                new DocumentFormat.OpenXml.Drawing.PictureLocks { NoChangeAspect = true }
            )
        ),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.BlipFill(
            new DocumentFormat.OpenXml.Drawing.Blip { Embed = "rId1" },
            new DocumentFormat.OpenXml.Drawing.Stretch(new DocumentFormat.OpenXml.Drawing.FillRectangle())
        ),
        new DocumentFormat.OpenXml.Drawing.Spreadsheet.ShapeProperties(
            new DocumentFormat.OpenXml.Drawing.Transform2D(
                new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                new DocumentFormat.OpenXml.Drawing.Extents { Cx = 0L, Cy = 0L }
            ),
            new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                new DocumentFormat.OpenXml.Drawing.AdjustValueList()
            ) { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }
        )
    ),
    new DocumentFormat.OpenXml.Drawing.Spreadsheet.ClientData()
);

drawingsPart.WorksheetDrawing.Append(anchor);
drawingsPart.WorksheetDrawing.Save();

worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet
{
    Id = workbookPart.GetIdOfPart(worksheetPart),
    SheetId = 1U,
    Name = "Sheet1"
});
workbookPart.Workbook.Save();

document.Dispose();
