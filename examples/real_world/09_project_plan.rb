# frozen_string_literal: true

require "date"
require "xlsxrb"

Xlsxrb.generate("09_project_plan.xlsx") do |wb|
  wb.sheet("Plan") do |sheet|
    sheet.row(%w[Phase Task Assignee Start End Status])
    sheet.row(["Phase 1", "Requirements", "Alice", Date.new(2026, 3, 1), Date.new(2026, 3, 15), "Done"])
    sheet.row(["Phase 2", "Implementation", "Bob", Date.new(2026, 3, 16), Date.new(2026, 4, 30), "In Progress"])
    sheet.row(["Phase 3", "Deployment", "Charlie", Date.new(2026, 5, 1), Date.new(2026, 5, 5), "Not Started"])
  end
end
