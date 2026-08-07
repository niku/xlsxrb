# frozen_string_literal: true

require "xlsxrb"
require "date"

Xlsxrb.generate("18_social_media_calendar.xlsx") do |wb|
  wb.style("header", font: { bold: true })
  wb.style("date") { |s| s.num_fmt("yyyy-mm-dd") }

  wb.sheet("Content Calendar") do |sheet|
    sheet.row(["Post Date", "Platform", "Content Topic", "Status", "Engagement"], styles: ["header"] * 5)
    sheet.column(0, width: 15)
    sheet.column(1, width: 15)
    sheet.column(2, width: 40)

    posts = [
      [Date.today + 1, "Twitter", "Product Launch Teaser", "Draft", 0],
      [Date.today + 2, "LinkedIn", "Company Culture Post", "Approved", 0],
      [Date.today + 3, "Instagram", "Behind the Scenes", "Planning", 0]
    ]

    posts.each do |post|
      sheet.row(post, styles: { 0 => "date" })
    end
  end
end
puts "18_social_media_calendar.xlsx generated"
