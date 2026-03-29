// Validates comment retention: XLSX should still contain comments.
var document = SpreadsheetDocument.Open(XlsxPath, false);
try
{
    var workbookPart = document.WorkbookPart ?? throw new Exception("WorkbookPart is missing.");
    var firstSheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault()
        ?? throw new Exception("Sheet is missing.");
    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(firstSheet.Id.Value);

    var commentsPart = worksheetPart.WorksheetCommentsPart
        ?? throw new Exception("WorksheetCommentsPart is missing after retention.");

    var comments = commentsPart.Comments ?? throw new Exception("Comments lost during retention.");
    var commentList = comments.Elements<CommentList>().FirstOrDefault()
        ?? throw new Exception("CommentList lost during retention.");

    var allComments = commentList.Elements<Comment>().ToList();
    if (allComments.Count < 2)
        throw new Exception($"Expected at least 2 comments after retention, got {allComments.Count}.");
}
finally
{
    document.Dispose();
}
