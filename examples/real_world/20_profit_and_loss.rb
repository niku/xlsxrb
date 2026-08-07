# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("20_profit_and_loss.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
  wb.style("bold_currency", font: { bold: true }) { |s| s.num_fmt("$#,##0.00") }

  wb.sheet("P&L") do |sheet|
    sheet.column(0, width: 20)
    (1..5).each { |i| sheet.column(i, width: 12) }

    sheet.row(%w[Category Q1 Q2 Q3 Q4 Total], styles: ["header"] * 6)

    curr_arr = { 1 => "currency", 2 => "currency", 3 => "currency", 4 => "currency", 5 => "currency" }
    bcurr_arr = { 1 => "bold_currency", 2 => "bold_currency", 3 => "bold_currency", 4 => "bold_currency", 5 => "bold_currency" }

    sheet.row(["Revenue", 50_000, 55_000, 60_000, 65_000, "=SUM(B2:E2)"], styles: curr_arr)
    sheet.row(["COGS", 20_000, 22_000, 25_000, 27_000, "=SUM(B3:E3)"], styles: curr_arr)
    sheet.row(["Gross Profit", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=F2-F3"], styles: bcurr_arr)

    sheet.row([])

    sheet.row(["Operating Exp", 15_000, 15_500, 16_000, 16_500, "=SUM(B6:E6)"], styles: curr_arr)
    sheet.row(["Net Income", "=B4-B6", "=C4-C6", "=D4-D6", "=E4-E6", "=F4-F6"], styles: bcurr_arr)
  end
end
puts "20_profit_and_loss.xlsx generated"
