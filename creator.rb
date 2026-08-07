# frozen_string_literal: true

scenarios = %w[
  21_to_do_list
  22_bug_tracker
  23_property_analysis
  24_crm
  25_travel_itinerary
  26_purchase_order
  27_bill_of_materials
  28_subscription_tracker
  29_fleet_management
  30_shift_schedule
]

scenarios.each do |scenario|
  File.write("examples/real_world/#{scenario}.rb", <<~RUBY)
    require "xlsxrb"

    Xlsxrb.generate("#{scenario}.xlsx") do |wb|
      wb.sheet("Sheet1") do |sheet|
        sheet.row(["ID", "Name", "Description", "Status"])
        sheet.row([1, "Item 1", "Description of Item 1", "Active"])
        sheet.row([2, "Item 2", "Description of Item 2", "Pending"])
      end
    end
  RUBY
end
