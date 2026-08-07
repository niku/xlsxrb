# frozen_string_literal: true

require "date"
require "xlsxrb"

Xlsxrb.generate("01_invoice.xlsx") do |wb|
  wb.sheet("Invoice") do |sheet|
    sheet.row(["Invoice #", "1001"])
    sheet.row(["Date", Date.today])
    sheet.row([])
    sheet.row(%w[Item Quantity Price Total])
    sheet.row(["Widget A", 2, 10.0, 20.0])
    sheet.row(["Widget B", 1, 15.0, 15.0])
    sheet.row([])
    sheet.row(["", "", "Total", 35.0])
  end
end
