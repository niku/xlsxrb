# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("04_gradebook.xlsx") do |wb|
  wb.sheet("Grades") do |sheet|
    sheet.row(%w[Student Math Science English Average])
    sheet.row(["Alice", 95, 90, 92, 92.3])
    sheet.row(["Bob", 80, 85, 88, 84.3])
    sheet.row(["Charlie", 70, 75, 80, 75.0])
  end
end
