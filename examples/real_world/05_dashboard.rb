# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("05_dashboard.xlsx") do |wb|
  wb.sheet("Dashboard") do |sheet|
    sheet.row(%w[Metric Value Target Status])
    sheet.row(["Active Users", 1500, 2000, "Needs Work"])
    sheet.row(["Monthly MRR", 50_000, 45_000, "Exceeded"])
    sheet.row(["Churn Rate", 0.05, 0.02, "High"])
  end
end
