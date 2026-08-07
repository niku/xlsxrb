# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("11_budget_planner.xlsx") do |wb|
  wb.style("header", font: { bold: true }, fill: { color: "CCCCCC", pattern: "solid" })
  wb.style("currency") { |s| s.num_fmt("$#,##0.00") }

  wb.sheet("Budget Planner") do |sheet|
    sheet.column(0, width: 20)
    sheet.column(1, width: 15)
    sheet.column(2, width: 15)
    sheet.row(%w[Category Estimated Actual Difference], styles: ["header"] * 4)

    data = [
      ["Housing", 1500, 1500],
      ["Food", 500, 600],
      ["Transport", 200, 150],
      ["Entertainment", 300, 350]
    ]

    data.each_with_index do |(cat, est, act), idx|
      row_num = idx + 2
      sheet.row([cat, est, act, "=B#{row_num}-C#{row_num}"], styles: { 1 => "currency", 2 => "currency", 3 => "currency" })
    end

    total_row = data.size + 2
    sheet.row(["Total", "=SUM(B2:B#{total_row - 1})", "=SUM(C2:C#{total_row - 1})", "=SUM(D2:D#{total_row - 1})"], styles: { 0 => "header", 1 => "currency", 2 => "currency", 3 => "currency" })
  end
end
puts "11_budget_planner.xlsx generated"
