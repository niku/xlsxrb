# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "borders.xlsx"

Xlsxrb.generate(output_path) do |w|
  w.style("thin") { |s| s.border_all(style: "thin", color: "FF000000") }
  w.style("medium") { |s| s.border_all(style: "medium", color: "FF000000") }
  w.style("thick") { |s| s.border_all(style: "thick", color: "FF000000") }
  w.style("hair") { |s| s.border_all(style: "hair", color: "FF000000") }
  w.style("dashed") { |s| s.border_all(style: "dashed", color: "FF000000") }
  w.style("medium_dashed") { |s| s.border_all(style: "mediumDashed", color: "FF000000") }
  w.style("dotted") { |s| s.border_all(style: "dotted", color: "FF000000") }
  w.style("double") { |s| s.border_all(style: "double", color: "FF000000") }
  w.style("dash_dot") { |s| s.border_all(style: "dashDot", color: "FF000000") }
  w.style("medium_dash_dot") { |s| s.border_all(style: "mediumDashDot", color: "FF000000") }
  w.style("dash_dot_dot") { |s| s.border_all(style: "dashDotDot", color: "FF000000") }
  w.style("slanted") { |s| s.border_all(style: "slantedDashDot", color: "FF000000") }
  w.style("diagonal") { |s| s.border_diagonal(style: "thin", color: "FF000000", up: true, down: true) }

  w.sheet("Borders") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.set_print_option(:grid_lines, true)

    s.row(["Border Style", "Cell Preview"])
    s.row(["Thin", "Thin Border"], styles: { 1 => "thin" })
    s.row(["Medium", "Medium Border"], styles: { 1 => "medium" })
    s.row(["Thick", "Thick Border"], styles: { 1 => "thick" })
    s.row(["Hair", "Hair Border"], styles: { 1 => "hair" })
    s.row(["Dashed", "Dashed Border"], styles: { 1 => "dashed" })
    s.row(["Medium Dashed", "Medium Dashed"], styles: { 1 => "medium_dashed" })
    s.row(["Dotted", "Dotted Border"], styles: { 1 => "dotted" })
    s.row(["Double", "Double Border"], styles: { 1 => "double" })
    s.row(["Dash-Dot", "Dash-Dot Border"], styles: { 1 => "dash_dot" })
    s.row(["Medium Dash-Dot", "Medium Dash-Dot"], styles: { 1 => "medium_dash_dot" })
    s.row(%w[Dash-Dot-Dot Dash-Dot-Dot], styles: { 1 => "dash_dot_dot" })
    s.row(["Slanted Dash-Dot", "Slanted Border"], styles: { 1 => "slanted" })
    s.row(["Diagonal (Cross)", "Diagonal Border"], styles: { 1 => "diagonal" })
  end
end

# Read check
puts "=== Read Validation ==="
workbook = Xlsxrb.read(output_path)
sheet = workbook.sheets.first
sheet.rows.each do |row|
  row_cells = row.cells.map do |c|
    xf = workbook.styles[:cell_xfs][c.style_index] if c.style_index
    border_id = xf ? xf[:border_id] : 0
    "#{c.ref}: #{c.value.inspect} (border_id: #{border_id})"
  end
  puts "Row #{row.index}: #{row_cells.join(", ")}"
end
