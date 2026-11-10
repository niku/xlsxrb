# frozen_string_literal: true

require "xlsxrb"

output_path = ARGV[0] || "borders.xlsx"

Xlsxrb.write(output_path) do |wb|
  wb.style("thin") { |s| s.border_all(style: "thin", color: "FF000000") }
  wb.style("medium") { |s| s.border_all(style: "medium", color: "FF000000") }
  wb.style("thick") { |s| s.border_all(style: "thick", color: "FF000000") }
  wb.style("hair") { |s| s.border_all(style: "hair", color: "FF000000") }
  wb.style("dashed") { |s| s.border_all(style: "dashed", color: "FF000000") }
  wb.style("medium_dashed") { |s| s.border_all(style: "mediumDashed", color: "FF000000") }
  wb.style("dotted") { |s| s.border_all(style: "dotted", color: "FF000000") }
  wb.style("double") { |s| s.border_all(style: "double", color: "FF000000") }
  wb.style("dash_dot") { |s| s.border_all(style: "dashDot", color: "FF000000") }
  wb.style("medium_dash_dot") { |s| s.border_all(style: "mediumDashDot", color: "FF000000") }
  wb.style("dash_dot_dot") { |s| s.border_all(style: "dashDotDot", color: "FF000000") }
  wb.style("slanted") { |s| s.border_all(style: "slantDashDot", color: "FF000000") }
  wb.style("diagonal") { |s| s.border_diagonal(style: "thin", color: "FF000000", up: true, down: true) }

  wb.sheet("Borders") do |s|
    s.column(0, width: 25)
    s.column(1, width: 25)
    s.print_options(grid_lines: true)

    s.row(["Border Style", "Cell Preview"])
    s.row(["Thin", "Thin Border"], cell_styles: { 1 => "thin" })
    s.row(["Medium", "Medium Border"], cell_styles: { 1 => "medium" })
    s.row(["Thick", "Thick Border"], cell_styles: { 1 => "thick" })
    s.row(["Hair", "Hair Border"], cell_styles: { 1 => "hair" })
    s.row(["Dashed", "Dashed Border"], cell_styles: { 1 => "dashed" })
    s.row(["Medium Dashed", "Medium Dashed"], cell_styles: { 1 => "medium_dashed" })
    s.row(["Dotted", "Dotted Border"], cell_styles: { 1 => "dotted" })
    s.row(["Double", "Double Border"], cell_styles: { 1 => "double" })
    s.row(["Dash-Dot", "Dash-Dot Border"], cell_styles: { 1 => "dash_dot" })
    s.row(["Medium Dash-Dot", "Medium Dash-Dot"], cell_styles: { 1 => "medium_dash_dot" })
    s.row(%w[Dash-Dot-Dot Dash-Dot-Dot], cell_styles: { 1 => "dash_dot_dot" })
    s.row(["Slanted Dash-Dot", "Slanted Border"], cell_styles: { 1 => "slanted" })
    s.row(["Diagonal (Cross)", "Diagonal Border"], cell_styles: { 1 => "diagonal" })
  end
end

# Read check
