var document = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);

try
{
    var workbookPart = document.AddWorkbookPart();
    workbookPart.Workbook = new Workbook();

    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    var sheetData = new SheetData(
        new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString, InlineString = new InlineString(new Text("hello")) })
    );
    worksheetPart.Worksheet = new Worksheet(sheetData);

    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
    sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });

    var extProps = document.AddExtendedFilePropertiesPart();
    extProps.Properties = new DocumentFormat.OpenXml.ExtendedProperties.Properties();
    extProps.Properties.Application = new DocumentFormat.OpenXml.ExtendedProperties.Application("SDK App");
    extProps.Properties.ApplicationVersion = new DocumentFormat.OpenXml.ExtendedProperties.ApplicationVersion("2.0.0");

    workbookPart.Workbook.Save();
    worksheetPart.Worksheet.Save();
    extProps.Properties.Save();
}
finally
{
    document.Dispose();
}
