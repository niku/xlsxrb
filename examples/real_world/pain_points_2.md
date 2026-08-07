# DSL Friction Points 2 (Scenarios 11-20)

1. **Inconsistent Style Builder API:** `wb.style` accepts a block for some properties like `num_fmt` (`wb.style("curr") { |s| s.num_fmt("...") }`) but lacks methods for others like `font` and `fill`, which must be passed as keyword arguments (e.g., `wb.style("header", font: { bold: true })`). Attempting to use `s.font(bold: true)` inside the block throws a `NoMethodError`. This inconsistency makes the DSL hard to guess and unintuitive.

2. **No Range Support in Styles Hash:** The `styles:` kwarg in `sheet.row` accepts a hash of column indices to style names (e.g., `styles: { 0 => "header", 1 => "currency" }`). However, providing a Range (e.g., `styles: { 0..4 => "header" }`) crashes the builder with an `Invalid column letter: 0..4` error. As a result, applying a style to multiple adjacent cells requires either providing a full array (`["header"] * 5`) or verbosely spelling out every index in the hash.
