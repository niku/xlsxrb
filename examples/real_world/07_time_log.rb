# frozen_string_literal: true

require "date"
require "xlsxrb"

Xlsxrb.generate("07_time_log.xlsx") do |wb|
  wb.sheet("TimeLog") do |sheet|
    sheet.row(%w[Date Project Task Hours Billable])
    sheet.row([Date.today, "Project Alpha", "Development", 4.5, true])
    sheet.row([Date.today, "Internal", "Meeting", 1.0, false])
    sheet.row([Date.today - 1, "Project Alpha", "Testing", 3.0, true])
  end
end
