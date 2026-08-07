# Xlsxrb DSL Friction Points

1. **Output Flexibility**: The `Xlsxrb.generate("file.xlsx")` method requires writing to a string/file. If I want to hold the document in memory and add sheets dynamically before rendering, I need to know about the difference between Streaming API (`generate`) and In-Memory API (`build`), which may not be immediately obvious.
2. **Formatting**: Using `sheet.row([...])` is very convenient, but applying formats (like currency for "Price", "Amount" or bolding headers) inline is not directly obvious from a basic use case. One might intuitively try `sheet.row([...], style: :bold)` or similar.
3. **Dates vs Strings**: Dates work well, but knowing how they are formatted in the resulting Excel file (e.g. "YYYY-MM-DD" vs "MM/DD/YYYY") is opaque without diving into the API documentation.
4. **Column widths**: The DSL supports `sheet.column(index, width: ...)` which is nice, but auto-fitting columns based on content isn't clearly exposed for a quick simple script.
