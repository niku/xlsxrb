// Creates an XLSX with comments on cells.
var doc = SpreadsheetDocument.Create(XlsxPath, SpreadsheetDocumentType.Workbook);
var workbookPart = doc.AddWorkbookPart();
workbookPart.Workbook = new Workbook();

var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
var sheetData = new SheetData(
    new Row(new Cell { CellReference = "A1", DataType = CellValues.InlineString,
        InlineString = new InlineString(new DocumentFormat.OpenXml.Spreadsheet.Text("data")) })
);
worksheetPart.Worksheet = new Worksheet(sheetData);

// Create comments part.
var commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>();
var comments = new Comments();
var authors = new Authors();
authors.Append(new Author("TestAuthor"));
authors.Append(new Author("SecondAuthor"));
comments.Append(authors);

var commentList = new CommentList();
var comment1 = new Comment { Reference = "A1", AuthorId = 0 };
var commentText1 = new CommentText();
var run1 = new DocumentFormat.OpenXml.Spreadsheet.Run();
run1.Append(new DocumentFormat.OpenXml.Spreadsheet.Text("First comment"));
commentText1.Append(run1);
comment1.Append(commentText1);
commentList.Append(comment1);

var comment2 = new Comment { Reference = "B2", AuthorId = 1 };
var commentText2 = new CommentText();
var run2 = new DocumentFormat.OpenXml.Spreadsheet.Run();
run2.Append(new DocumentFormat.OpenXml.Spreadsheet.Text("Second comment"));
commentText2.Append(run2);
comment2.Append(commentText2);
commentList.Append(comment2);

comments.Append(commentList);
commentsPart.Comments = comments;
commentsPart.Comments.Save();

worksheetPart.Worksheet.Save();

var sheets = workbookPart.Workbook.AppendChild(new Sheets());
sheets.Append(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });
workbookPart.Workbook.Save();
doc.Dispose();
