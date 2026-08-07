# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("14_recipe_cost_calculator.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("currency") { |s| s.num_fmt("$#,##0.00") }

  wb.sheet("Recipe Cost") do |sheet|
    sheet.row(["Ingredient", "Quantity", "Unit", "Cost per Unit", "Total Cost"], styles: ["header"] * 5)

    ingredients = [
      ["Flour", 2, "kg", 1.50],
      ["Sugar", 0.5, "kg", 2.00],
      ["Eggs", 12, "pcs", 0.20],
      ["Butter", 0.25, "kg", 8.00]
    ]

    ingredients.each_with_index do |(ing, qty, unit, cost), idx|
      r = idx + 2
      sheet.row([ing, qty, unit, cost, "=B#{r}*D#{r}"], styles: { 3 => "currency", 4 => "currency" })
    end

    r = ingredients.size + 2
    sheet.row(["Total Recipe Cost", "", "", "", "=SUM(E2:E#{r - 1})"], styles: { 0 => "header", 4 => "currency" })
  end
end
puts "14_recipe_cost_calculator.xlsx generated"
