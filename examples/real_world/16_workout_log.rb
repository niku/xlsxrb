# frozen_string_literal: true

require "xlsxrb"
require "date"

Xlsxrb.generate("16_workout_log.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("date") { |s| s.num_fmt("yyyy-mm-dd") }

  wb.sheet("Workouts") do |sheet|
    sheet.row(["Date", "Type", "Duration (min)", "Avg HR", "Calories Burned"], styles: ["header"] * 5)

    logs = [
      [Date.today - 2, "Running", 45, 155, 450],
      [Date.today - 1, "Cycling", 60, 140, 500],
      [Date.today, "Swimming", 30, 130, 300]
    ]

    logs.each do |log|
      sheet.row(log, styles: { 0 => "date" })
    end
  end
end
puts "16_workout_log.xlsx generated"
