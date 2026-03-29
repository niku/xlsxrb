var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text("hello")) })
);
var printOpts = new PrintOptions { GridLines = true, HorizontalCentered = true };
var pageMargins = new PageMargins { Left = 0.5, Right = 0.5, Top = 1.0, Bottom = 1.0, Header = 0.3, Footer = 0.3 };
var pageSetup = new PageSetup { Orientation = OrientationValues.Portrait, PaperSize = (uint)9, Scale = (uint)80 };
var headerFooter = new HeaderFooter(
    new OddHeader { Text = "&LLeft Header" },
    new OddFooter { Text = "&RRight Footer" }
);
var rowBreaks = new RowBreaks(
    new Break { Id = (uint)5, Max = (uint)16383, ManualPageBreak = true },
    new Break { Id = (uint)15, Max = (uint)16383, ManualPageBreak = true }
) { Count = 2, ManualBreakCount = 2 };

worksheetPart.Worksheet = new Worksheet(sheetData, printOpts, pageMargins, pageSetup, headerFooter, rowBreaks);
worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
