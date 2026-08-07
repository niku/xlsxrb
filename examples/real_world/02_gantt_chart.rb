# frozen_string_literal: true

require "date"
require "xlsxrb"

Xlsxrb.generate("02_gantt_chart.xlsx") do |wb|
  wb.sheet("Gantt") do |sheet|
    sheet.row(["Task", "Start Date", "End Date", "Duration"])
    sheet.row(["Design", Date.new(2026, 1, 1), Date.new(2026, 1, 10), 9])
    sheet.row(["Develop", Date.new(2026, 1, 11), Date.new(2026, 2, 20), 40])
    sheet.row(["Test", Date.new(2026, 2, 21), Date.new(2026, 2, 28), 7])
  end
end
