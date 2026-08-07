# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("24_crm.xlsx") do |wb|
  wb.sheet("Sheet1") do |sheet|
    sheet.row(%w[ID Name Description Status])
    sheet.row([1, "Item 1", "Description of Item 1", "Active"])
    sheet.row([2, "Item 2", "Description of Item 2", "Pending"])
  end
end
