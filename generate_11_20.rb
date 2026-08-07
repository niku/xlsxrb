# frozen_string_literal: true

require "fileutils"

FileUtils.mkdir_p("examples/real_world")

files = {
  "examples/real_world/11_budget_planner.rb" => <<~'RUBY',
    # frozen_string_literal: true
    require "xlsxrb"

    Xlsxrb.generate("11_budget_planner.xlsx") do |wb|
      wb.style("header", font: { bold: true }, fill: { color: "CCCCCC", pattern: "solid" })
      wb.style("currency") { |s| s.num_fmt("$#,##0.00") }

      wb.sheet("Budget Planner") do |sheet|
        sheet.column(0, width: 20)
        sheet.column(1, width: 15)
        sheet.column(2, width: 15)
        sheet.row(["Category", "Estimated", "Actual", "Difference"], styles: ["header"] * 4)

        data = [
          ["Housing", 1500, 1500],
          ["Food", 500, 600],
          ["Transport", 200, 150],
          ["Entertainment", 300, 350]
        ]

        data.each_with_index do |(cat, est, act), idx|
          row_num = idx + 2
          sheet.row([cat, est, act, "=B#{row_num}-C#{row_num}"], styles: {1 => "currency", 2 => "currency", 3 => "currency"})
        end

        total_row = data.size + 2
        sheet.row(["Total", "=SUM(B2:B#{total_row-1})", "=SUM(C2:C#{total_row-1})", "=SUM(D2:D#{total_row-1})"], styles: {0 => "header", 1 => "currency", 2 => "currency", 3 => "currency"})
      end
    end
    puts "11_budget_planner.xlsx generated"
  RUBY

  "examples/real_world/12_sales_pipeline.rb" => <<~'RUBY',
    # frozen_string_literal: true
    require "xlsxrb"

    Xlsxrb.generate("12_sales_pipeline.xlsx") do |wb|
      wb.style("header", font: { bold: true })
      wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
      wb.style("percent") { |s| s.num_fmt("0%") }

      wb.sheet("Sales Pipeline") do |sheet|
        sheet.row(["Client", "Deal Value", "Probability", "Expected Value", "Stage"], styles: ["header"] * 5)
        sheet.column(0, width: 25)
        sheet.column(1, width: 15)
        sheet.column(2, width: 15)
        sheet.column(3, width: 15)
        sheet.column(4, width: 15)

        deals = [
          ["Acme Corp", 10000, 0.5, "Proposal"],
          ["Globex", 50000, 0.2, "Lead"],
          ["Soylent", 5000, 0.9, "Closing"]
        ]

        deals.each_with_index do |(client, val, prob, stage), idx|
          r = idx + 2
          sheet.row([client, val, prob, "=B#{r}*C#{r}", stage], styles: {1 => "currency", 2 => "percent", 3 => "currency"})
        end
      end
    end
    puts "12_sales_pipeline.xlsx generated"
  RUBY

  "examples/real_world/13_fitness_tracker.rb" => <<~'RUBY',
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
          sheet.row([date, ex, sets, reps, weight, "=C#{r}*D#{r}*E#{r}"], styles: {0 => "date"})
        end
      end
    end
    puts "13_fitness_tracker.xlsx generated"
  RUBY

  "examples/real_world/14_recipe_cost_calculator.rb" => <<~'RUBY',
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
          sheet.row([ing, qty, unit, cost, "=B#{r}*D#{r}"], styles: {3 => "currency", 4 => "currency"})
        end

        r = ingredients.size + 2
        sheet.row(["Total Recipe Cost", "", "", "", "=SUM(E2:E#{r-1})"], styles: {0 => "header", 4 => "currency"})
      end
    end
    puts "14_recipe_cost_calculator.xlsx generated"
  RUBY

  "examples/real_world/15_event_guest_list.rb" => <<~RUBY,
    # frozen_string_literal: true
    require "xlsxrb"

    Xlsxrb.generate("15_event_guest_list.xlsx") do |wb|
      wb.style("header", font: { bold: true })
    #{"  "}
      wb.sheet("Guest List") do |sheet|
        sheet.row(["Name", "RSVP Status", "Plus One", "Dietary Requirements"], styles: ["header"] * 4)
    #{"    "}
        guests = [
          ["Alice Smith", true, false, "Vegan"],
          ["Bob Jones", false, false, "None"],
          ["Charlie Brown", true, true, "Gluten Free"]
        ]
    #{"    "}
        guests.each do |guest|
          sheet.row(guest)
        end
      end
    end
    puts "15_event_guest_list.xlsx generated"
  RUBY

  "examples/real_world/16_workout_log.rb" => <<~RUBY,
    # frozen_string_literal: true
    require "xlsxrb"
    require "date"

    Xlsxrb.generate("16_workout_log.xlsx") do |wb|
      wb.style("header", font: { bold: true })
      wb.style("date") { |s| s.num_fmt("yyyy-mm-dd") }
    #{"  "}
      wb.sheet("Workouts") do |sheet|
        sheet.row(["Date", "Type", "Duration (min)", "Avg HR", "Calories Burned"], styles: ["header"] * 5)
    #{"    "}
        logs = [
          [Date.today - 2, "Running", 45, 155, 450],
          [Date.today - 1, "Cycling", 60, 140, 500],
          [Date.today, "Swimming", 30, 130, 300]
        ]
    #{"    "}
        logs.each do |log|
          sheet.row(log, styles: {0 => "date"})
        end
      end
    end
    puts "16_workout_log.xlsx generated"
  RUBY

  "examples/real_world/17_mortgage_amortization.rb" => <<~RUBY,
    # frozen_string_literal: true
    require "xlsxrb"

    Xlsxrb.generate("17_mortgage_amortization.xlsx") do |wb|
      wb.style("header", font: { bold: true })
      wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
      wb.style("bold_currency", font: { bold: true }) { |s| s.num_fmt("$#,##0.00") }
    #{"  "}
      wb.sheet("Amortization") do |sheet|
        sheet.column(0, width: 10)
        (1..5).each { |i| sheet.column(i, width: 18) }
    #{"    "}
        sheet.row(["Loan Amount", 300000], styles: {1 => "bold_currency"})
        sheet.row(["Interest Rate", 0.05])
        sheet.row(["Months", 360])
        sheet.row([])
    #{"    "}
        sheet.row(["Month", "Beg. Balance", "Payment", "Interest", "Principal", "End Balance"], styles: ["header"] * 6)
    #{"    "}
        # Simplified first few months
        sheet.row([1, "=B1", 1610.46, "=B6*(B2/12)", "=C6-D6", "=B6-E6"], styles: {1 => "currency", 2 => "currency", 3 => "currency", 4 => "currency", 5 => "currency"})
        sheet.row([2, "=F6", 1610.46, "=B7*(B2/12)", "=C7-D7", "=B7-E7"], styles: {1 => "currency", 2 => "currency", 3 => "currency", 4 => "currency", 5 => "currency"})
      end
    end
    puts "17_mortgage_amortization.xlsx generated"
  RUBY

  "examples/real_world/18_social_media_calendar.rb" => <<~RUBY,
    # frozen_string_literal: true
    require "xlsxrb"
    require "date"

    Xlsxrb.generate("18_social_media_calendar.xlsx") do |wb|
      wb.style("header", font: { bold: true })
      wb.style("date") { |s| s.num_fmt("yyyy-mm-dd") }
    #{"  "}
      wb.sheet("Content Calendar") do |sheet|
        sheet.row(["Post Date", "Platform", "Content Topic", "Status", "Engagement"], styles: ["header"] * 5)
        sheet.column(0, width: 15)
        sheet.column(1, width: 15)
        sheet.column(2, width: 40)
    #{"    "}
        posts = [
          [Date.today + 1, "Twitter", "Product Launch Teaser", "Draft", 0],
          [Date.today + 2, "LinkedIn", "Company Culture Post", "Approved", 0],
          [Date.today + 3, "Instagram", "Behind the Scenes", "Planning", 0]
        ]
    #{"    "}
        posts.each do |post|
          sheet.row(post, styles: {0 => "date"})
        end
      end
    end
    puts "18_social_media_calendar.xlsx generated"
  RUBY

  "examples/real_world/19_maintenance_schedule.rb" => <<~'RUBY',
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
          sheet.row([eq, last, freq, "=B#{r}+C#{r}"], styles: {1 => "date", 3 => "date"})
        end
      end
    end
    puts "19_maintenance_schedule.xlsx generated"
  RUBY

  "examples/real_world/20_profit_and_loss.rb" => <<~RUBY
    # frozen_string_literal: true
    require "xlsxrb"

    Xlsxrb.generate("20_profit_and_loss.xlsx") do |wb|
      wb.style("header", font: { bold: true })
      wb.style("currency") { |s| s.num_fmt("$#,##0.00") }
      wb.style("bold_currency", font: { bold: true }) { |s| s.num_fmt("$#,##0.00") }
    #{"  "}
      wb.sheet("P&L") do |sheet|
        sheet.column(0, width: 20)
        (1..5).each { |i| sheet.column(i, width: 12) }
    #{"    "}
        sheet.row(["Category", "Q1", "Q2", "Q3", "Q4", "Total"], styles: ["header"] * 6)
    #{"    "}
        curr_arr = {1 => "currency", 2 => "currency", 3 => "currency", 4 => "currency", 5 => "currency"}
        bcurr_arr = {1 => "bold_currency", 2 => "bold_currency", 3 => "bold_currency", 4 => "bold_currency", 5 => "bold_currency"}

        sheet.row(["Revenue", 50000, 55000, 60000, 65000, "=SUM(B2:E2)"], styles: curr_arr)
        sheet.row(["COGS", 20000, 22000, 25000, 27000, "=SUM(B3:E3)"], styles: curr_arr)
        sheet.row(["Gross Profit", "=B2-B3", "=C2-C3", "=D2-D3", "=E2-E3", "=F2-F3"], styles: bcurr_arr)
    #{"    "}
        sheet.row([])
    #{"    "}
        sheet.row(["Operating Exp", 15000, 15500, 16000, 16500, "=SUM(B6:E6)"], styles: curr_arr)
        sheet.row(["Net Income", "=B4-B6", "=C4-C6", "=D4-D6", "=E4-E6", "=F4-F6"], styles: bcurr_arr)
      end
    end
    puts "20_profit_and_loss.xlsx generated"
  RUBY
}

files.each do |path, content|
  File.write(path, content)
  puts "Wrote #{path}"
end
