# frozen_string_literal: true

require "xlsxrb"
require "date"

Xlsxrb.generate("13_fitness_tracker.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("date") { |s| s.num_fmt("yyyy-mm-dd") }

  wb.sheet("Fitness Tracker") do |sheet|
    sheet.row(["Date", "Exercise", "Sets", "Reps", "Weight (kg)", "Volume"], styles: ["header"] * 6)

    logs = [
      [Date.today, "Squat", 3, 5, 100],
      [Date.today, "Bench Press", 3, 5, 80],
      [Date.today, "Deadlift", 1, 5, 120]
    ]

    logs.each_with_index do |(date, ex, sets, reps, weight), idx|
      r = idx + 2
      sheet.row([date, ex, sets, reps, weight, "=C#{r}*D#{r}*E#{r}"], styles: { 0 => "date" })
    end
  end
end
puts "13_fitness_tracker.xlsx generated"
