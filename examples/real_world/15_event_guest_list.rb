# frozen_string_literal: true

require "xlsxrb"

Xlsxrb.generate("15_event_guest_list.xlsx") do |wb|
  wb.style("header", font: { bold: true })

  wb.sheet("Guest List") do |sheet|
    sheet.row(["Name", "RSVP Status", "Plus One", "Dietary Requirements"], styles: ["header"] * 4)

    guests = [
      ["Alice Smith", true, false, "Vegan"],
      ["Bob Jones", false, false, "None"],
      ["Charlie Brown", true, true, "Gluten Free"]
    ]

    guests.each do |guest|
      sheet.row(guest)
    end
  end
end
puts "15_event_guest_list.xlsx generated"
