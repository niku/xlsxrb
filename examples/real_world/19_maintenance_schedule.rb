# frozen_string_literal: true

require "xlsxrb"
require "date"

Xlsxrb.generate("19_maintenance_schedule.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("date") { |s| s.num_fmt("yyyy-mm-dd") }

  wb.sheet("Schedule") do |sheet|
    sheet.row(["Equipment", "Last Maintenance", "Frequency (Days)", "Next Maintenance"], styles: ["header"] * 4)
    sheet.column(0, width: 25)
    sheet.column(1, width: 18)
    sheet.column(2, width: 18)
    sheet.column(3, width: 18)

    items = [
      ["HVAC System", Date.today - 100, 180],
      ["Forklift A", Date.today - 30, 90],
      ["Conveyor Belt", Date.today - 10, 30]
    ]

    items.each_with_index do |(eq, last, freq), idx|
      r = idx + 2
      sheet.row([eq, last, freq, "=B#{r}+C#{r}"], styles: { 1 => "date", 3 => "date" })
    end
  end
end
puts "19_maintenance_schedule.xlsx generated"
