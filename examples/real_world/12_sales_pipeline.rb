# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("12_sales_pipeline.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
  wb.style("percent") { |s| s.num_fmt("0%") }

  wb.sheet("Sales Pipeline") do |sheet|
    sheet.row(["Client", "Deal Value", "Probability", "Expected Value", "Stage"], styles: ["header"] * 5)
    sheet.column(0, width: 25)
    sheet.column(1, width: 15)
    sheet.column(2, width: 15)
    sheet.column(3, width: 15)
    sheet.column(4, width: 15)

    deals = [
      ["Acme Corp", 10_000, 0.5, "Proposal"],
      ["Globex", 50_000, 0.2, "Lead"],
      ["Soylent", 5000, 0.9, "Closing"]
    ]

    deals.each_with_index do |(client, val, prob, stage), idx|
      r = idx + 2
      sheet.row([client, val, prob, "=B#{r}*C#{r}", stage], styles: { 1 => "currency", 2 => "percent", 3 => "currency" })
    end
  end
end
puts "12_sales_pipeline.xlsx generated"
