# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("17_mortgage_amortization.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
  wb.style("bold_currency", font: { bold: true }) { |s| s.num_fmt("$#,##0.00") }

  wb.sheet("Amortization") do |sheet|
    sheet.column(0, width: 10)
    (1..5).each { |i| sheet.column(i, width: 18) }

    sheet.row(["Loan Amount", 300_000], styles: { 1 => "bold_currency" })
    sheet.row(["Interest Rate", 0.05])
    sheet.row(["Months", 360])
    sheet.row([])

    sheet.row(["Month", "Beg. Balance", "Payment", "Interest", "Principal", "End Balance"], styles: ["header"] * 6)

    # Simplified first few months
    sheet.row([1, "=B1", 1610.46, "=B6*(B2/12)", "=C6-D6", "=B6-E6"], styles: { 1 => "currency", 2 => "currency", 3 => "currency", 4 => "currency", 5 => "currency" })
    sheet.row([2, "=F6", 1610.46, "=B7*(B2/12)", "=C7-D7", "=B7-E7"], styles: { 1 => "currency", 2 => "currency", 3 => "currency", 4 => "currency", 5 => "currency" })
  end
end
puts "17_mortgage_amortization.xlsx generated"
