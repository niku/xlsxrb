# frozen_string_literal: true

require "date"
require "xlsxrb"

Xlsxrb.generate("08_expense_report.xlsx") do |wb|
  wb.sheet("Expenses") do |sheet|
    sheet.row(%w[Date Category Description Amount Approved])
    sheet.row([Date.today, "Travel", "Flight to NY", 450.0, true])
    sheet.row([Date.today, "Meals", "Client Dinner", 120.5, false])
    sheet.row([Date.today - 2, "Office", "Supplies", 35.0, true])
  end
end
