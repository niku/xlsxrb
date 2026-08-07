# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("03_financial_statement.xlsx") do |wb|
  wb.sheet("Financials") do |sheet|
    sheet.row(["Account", "Q1", "Q2", "Q3", "Q4", "Year Total"])
    sheet.row(["Revenue", 10_000, 12_000, 15_000, 20_000, 57_000])
    sheet.row(["COGS", -2000, -2500, -3000, -4000, -11_500])
    sheet.row(["Gross Margin", 8000, 9500, 12_000, 16_000, 45_500])
  end
end
