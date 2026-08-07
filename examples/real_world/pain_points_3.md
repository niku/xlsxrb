# DSL Friction Points 3

1. **Lack of column styling options:** The API is very simple (`sheet.row(...)`), which makes it difficult to specify column widths or row heights easily without dropping down to a more complex object model.
2. **Type conversion:** Explicit type conversion is mostly implicit in Ruby arrays. There's no easy way to enforce a cell format (e.g. force text for numbers with leading zeros) in a simple array pass.
3. **Complex layouts:** Creating multi-header tables or merging cells is not immediately obvious in the DSL block when just yielding rows.
4. **Data vs Structure:** Hardcoding data in rows via `sheet.row` works for small scripts but scales poorly for larger applications where data and presentation should be separated.
