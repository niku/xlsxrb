var document = SpreadsheetDocument.Open(XlsxPath, false);

try
{
    var worksheetPart = document.WorkbookPart.WorksheetParts.First();
    var ws = worksheetPart.Worksheet;
    var autoFilter = ws.GetFirstChild<AutoFilter>();
    if (autoFilter == null) throw new Exception("AutoFilter element not found.");
    var filterColumns = autoFilter.Elements<FilterColumn>().ToList();
    if (filterColumns.Count < 1) throw new Exception($"Expected at least 1 filter column but got {filterColumns.Count}.");

    var fc0 = filterColumns[0];
    if (fc0.HiddenButton == null || fc0.HiddenButton.Value != true)
        throw new Exception($"Expected hiddenButton=true but got {fc0.HiddenButton?.Value}.");
    if (fc0.ShowButton == null || fc0.ShowButton.Value != false)
        throw new Exception($"Expected showButton=false but got {fc0.ShowButton?.Value}.");

    var filters = fc0.GetFirstChild<Filters>();
    if (filters == null) throw new Exception("Filters element not found.");
    if (filters.CalendarType == null || filters.CalendarType.Value != CalendarValues.Gregorian)
        throw new Exception($"Expected calendarType=gregorian but got {filters.CalendarType?.Value}.");

    var dateGroupItems = filters.Elements<DateGroupItem>().ToList();
    if (dateGroupItems.Count != 1) throw new Exception($"Expected 1 dateGroupItem but got {dateGroupItems.Count}.");
    if (dateGroupItems[0].DateTimeGrouping == null || dateGroupItems[0].DateTimeGrouping.Value != DateTimeGroupingValues.Year)
        throw new Exception($"Expected dateTimeGrouping=year but got {dateGroupItems[0].DateTimeGrouping?.Value}.");
    if (dateGroupItems[0].Year == null || dateGroupItems[0].Year.Value != 2024)
        throw new Exception($"Expected year=2024 but got {dateGroupItems[0].Year?.Value}.");
}
finally
{
    document.Dispose();
}
