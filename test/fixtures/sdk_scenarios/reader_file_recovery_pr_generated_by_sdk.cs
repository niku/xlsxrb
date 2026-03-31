// Creates an XLSX with fileRecoveryPr in the workbook.
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
worksheetPart.Worksheet = new Worksheet(new SheetData());
worksheetPart.Worksheet.Save();

var sheetsEl = workbookPart.Workbook.AppendChild(new Sheets());
sheetsEl.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });

workbookPart.Workbook.AppendChild(new FileRecoveryProperties { AutoRecover = false, CrashSave = true });
workbookPart.Workbook.Save();
doc.Dispose();

Console.Error.WriteLine("SCENARIO_PASS");
