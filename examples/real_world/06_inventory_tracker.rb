# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("06_inventory_tracker.xlsx") do |wb|
  wb.sheet("Inventory") do |sheet|
    sheet.row(["Item ID", "Name", "Stock", "Reorder Level", "Unit Price"])
    sheet.row(["ITM-001", "Laptop", 45, 10, 1200.0])
    sheet.row(["ITM-002", "Mouse", 150, 50, 25.0])
    sheet.row(["ITM-003", "Keyboard", 8, 20, 45.0])
  end
end
