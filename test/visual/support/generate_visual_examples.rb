# frozen_string_literal: true

require "fileutils"

module Xlsxrb
  module Visual
    class ExampleGenerator
      FILES = {
        "cell_numbers.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cell_numbers.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("currency") { |s| s.num_fmt("$#,##0.00") }
            w.add_style("percent") { |s| s.num_fmt("0.0%") }
            w.add_sheet("Numbers") do |s|
              s.add_row(["Format", "Value"])
              s.add_row(["Integer", 12345])
              s.add_row(["Float", 123.456])
              s.add_row(["Currency", 1234.5], styles: { 1 => "currency" })
              s.add_row(["Percentage", 0.85], styles: { 1 => "percent" })
            end
          end
        RUBY

        "cell_dates.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          require "date"
          output_path = ARGV[0] || "cell_dates.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("custom_date") { |s| s.num_fmt("yyyy-mm-dd") }
            w.add_sheet("Dates") do |s|
              s.add_row(["Format", "Date Value"])
              s.add_row(["Default Date", Date.new(2026, 7, 1)])
              s.add_row(["Formatted Date", Date.new(2026, 12, 25)], styles: { 1 => "custom_date" })
            end
          end
        RUBY

        "cell_times.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          require "time"
          output_path = ARGV[0] || "cell_times.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("time_fmt") { |s| s.num_fmt("hh:mm:ss") }
            w.add_sheet("Times") do |s|
              s.add_row(["Format", "Time Value"])
              s.add_row(["DateTime", Time.new(2026, 7, 1, 12, 34, 56)])
              s.add_row(["Time Only", Time.new(2026, 7, 1, 9, 15, 0)], styles: { 1 => "time_fmt" })
            end
          end
        RUBY

        "cell_booleans.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cell_booleans.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Booleans") do |s|
              s.add_row(["Label", "Boolean Value"])
              s.add_row(["Is Active", true])
              s.add_row(["Is Pending", false])
            end
          end
        RUBY

        "cell_rich_text.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cell_rich_text.xlsx"
          Xlsxrb.generate(output_path) do |w|
            rt = Xlsxrb::Elements::RichText.new(runs: [
              { text: "Normal " },
              { text: "Bold ", font: { bold: true, color: "FFC00000" } },
              { text: "Italic", font: { italic: true, sz: 14 } }
            ])
            w.add_sheet("Rich Text") do |s|
              s.add_row(["Format", "Value"])
              s.add_row(["Rich Text Cell", rt])
            end
          end
        RUBY

        "cell_formulas.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cell_formulas.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Formulas") do |s|
              s.add_row(["Item", "Value"])
              s.add_row(["A", 10])
              s.add_row(["B", 20])
              s.add_row(["SUM", "=SUM(B2:B3)"])
              s.add_row(["AVERAGE", "=AVERAGE(B2:B3)"])
            end
          end
        RUBY

        "font_families.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_families.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("arial") { |s| s.font_name("Arial") }
            w.add_style("times") { |s| s.font_name("Times New Roman") }
            w.add_style("courier") { |s| s.font_name("Courier New") }
            w.add_sheet("Fonts") do |s|
              s.add_row(["Font Family", "Preview"])
              s.add_row(["Arial", "Hello Arial"], styles: { 1 => "arial" })
              s.add_row(["Times New Roman", "Hello Times"], styles: { 1 => "times" })
              s.add_row(["Courier New", "Hello Courier"], styles: { 1 => "courier" })
            end
          end
        RUBY

        "font_sizes.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_sizes.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("size_10") { |s| s.size(10) }
            w.add_style("size_16") { |s| s.size(16) }
            w.add_style("size_24") { |s| s.size(24) }
            w.add_sheet("Font Sizes") do |s|
              s.add_row(["Size", "Text"])
              s.add_row(["10pt", "Small Text"], styles: { 1 => "size_10" })
              s.add_row(["16pt", "Medium Text"], styles: { 1 => "size_16" })
              s.add_row(["24pt", "Large Text"], styles: { 1 => "size_24" })
            end
          end
        RUBY

        "font_colors.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_colors.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("red") { |s| s.font_color("FFC00000") }
            w.add_style("green") { |s| s.font_color("FF00B050") }
            w.add_style("blue") { |s| s.font_color("FF0070C0") }
            w.add_sheet("Colors") do |s|
              s.add_row(["Color", "Preview"])
              s.add_row(["Red", "Red Text"], styles: { 1 => "red" })
              s.add_row(["Green", "Green Text"], styles: { 1 => "green" })
              s.add_row(["Blue", "Blue Text"], styles: { 1 => "blue" })
            end
          end
        RUBY

        "font_styles.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_styles.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("bold") { |s| s.bold }
            w.add_style("italic") { |s| s.italic }
            w.add_style("underline") { |s| s.underline }
            w.add_style("strike") { |s| s.strike }
            w.add_sheet("Styles") do |s|
              s.add_row(["Style", "Preview"])
              s.add_row(["Bold", "Bold Text"], styles: { 1 => "bold" })
              s.add_row(["Italic", "Italic Text"], styles: { 1 => "italic" })
              s.add_row(["Underline", "Underlined Text"], styles: { 1 => "underline" })
              s.add_row(["Strike-through", "Struck Text"], styles: { 1 => "strike" })
            end
          end
        RUBY

        "font_vertical_align.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_vertical_align.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("super") { |s| s.vert_align("superscript") }
            w.add_style("sub") { |s| s.vert_align("subscript") }
            w.add_sheet("Vertical Align") do |s|
              s.add_row(["Format", "Text"])
              s.add_row(["Superscript", "x2"], styles: { 1 => "super" })
              s.add_row(["Subscript", "H2O"], styles: { 1 => "sub" })
            end
          end
        RUBY

        "border_thin.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_thin.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("thin") { |s| s.border_left(style: "thin").border_right(style: "thin").border_top(style: "thin").border_bottom(style: "thin") }
            w.add_sheet("Thin Borders") do |s|
              s.add_row(["Normal Cell", "Thin Border Cell"], styles: { 1 => "thin" })
            end
          end
        RUBY

        "border_medium.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_medium.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("medium") { |s| s.border_left(style: "medium").border_right(style: "medium").border_top(style: "medium").border_bottom(style: "medium") }
            w.add_sheet("Medium Borders") do |s|
              s.add_row(["Normal Cell", "Medium Border Cell"], styles: { 1 => "medium" })
            end
          end
        RUBY

        "border_dashed.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_dashed.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("dashed") { |s| s.border_left(style: "dashed").border_right(style: "dashed").border_top(style: "dashed").border_bottom(style: "dashed") }
            w.add_sheet("Dashed Borders") do |s|
              s.add_row(["Normal Cell", "Dashed Border Cell"], styles: { 1 => "dashed" })
            end
          end
        RUBY

        "border_double.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_double.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("double") { |s| s.border_left(style: "double").border_right(style: "double").border_top(style: "double").border_bottom(style: "double") }
            w.add_sheet("Double Borders") do |s|
              s.add_row(["Normal Cell", "Double Border Cell"], styles: { 1 => "double" })
            end
          end
        RUBY

        "border_slanted.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_slanted.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("slanted") { |s| s.border_left(style: "slantedDashDot").border_right(style: "slantedDashDot").border_top(style: "slantedDashDot").border_bottom(style: "slantedDashDot") }
            w.add_sheet("Slanted Borders") do |s|
              s.add_row(["Normal Cell", "Slanted Border Cell"], styles: { 1 => "slanted" })
            end
          end
        RUBY

        "fill_solid_colors.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "fill_solid_colors.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("red_fill") { |s| s.fill_color("FFFFC7CE") }
            w.add_style("green_fill") { |s| s.fill_color("FFC6EFCE") }
            w.add_sheet("Fills") do |s|
              s.add_row(["Color", "Preview"])
              s.add_row(["Red", "Red Fill"], styles: { 1 => "red_fill" })
              s.add_row(["Green", "Green Fill"], styles: { 1 => "green_fill" })
            end
          end
        RUBY

        "fill_patterns.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "fill_patterns.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("dark_gray") { |s| s.fill(pattern: "darkGray", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
            w.add_style("grid_fill") { |s| s.fill(pattern: "darkGrid", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
            w.add_sheet("Patterns") do |s|
              s.add_row(["Pattern", "Preview"])
              s.add_row(["Dark Gray", "Pattern Fill"], styles: { 1 => "dark_gray" })
              s.add_row(["Dark Grid", "Grid Fill"], styles: { 1 => "grid_fill" })
            end
          end
        RUBY

        "fill_gradients.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "fill_gradients.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("gradient") do |style|
              style.fill_gradient(type: "linear", degree: 45, stops: [
                { position: 0, color: "FFFFFFFF" },
                { position: 1, color: "FF4F81BD" }
              ])
            end
            w.add_sheet("Gradients") do |s|
              s.add_row(["Normal Cell", "Gradient Cell"], styles: { 1 => "gradient" })
            end
          end
        RUBY

        "align_horizontal.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_horizontal.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("left") { |s| s.align_horizontal("left") }
            w.add_style("center") { |s| s.align_horizontal("center") }
            w.add_style("right") { |s| s.align_horizontal("right") }
            w.add_sheet("Alignment") do |s|
              s.set_print_option(:grid_lines, true)
              s.set_column(0, width: 20)
              s.set_column(1, width: 20)
              s.set_column(2, width: 20)
              s.add_row(["Left", "Center", "Right"], styles: { 0 => "left", 1 => "center", 2 => "right" })
            end
          end
        RUBY

        "align_vertical.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_vertical.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("top") { |s| s.align_vertical("top") }
            w.add_style("center") { |s| s.align_vertical("center") }
            w.add_style("bottom") { |s| s.align_vertical("bottom") }
            w.add_sheet("Vertical Alignment") do |s|
              s.set_print_option(:grid_lines, true)
              s.add_row(["Top", "Center", "Bottom"], styles: { 0 => "top", 1 => "center", 2 => "bottom" }, height: 40)
            end
          end
        RUBY

        "align_text_rotation.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_text_rotation.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("rot_45") { |s| s.text_rotation(45) }
            w.add_style("rot_90") { |s| s.text_rotation(90) }
            w.add_sheet("Rotation") do |s|
              s.set_print_option(:grid_lines, true)
              s.add_row(["Rotated 45", "Rotated 90"], styles: { 0 => "rot_45", 1 => "rot_90" }, height: 50)
            end
          end
        RUBY

        "align_text_wrap.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_text_wrap.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("wrap") { |s| s.wrap_text }
            w.add_sheet("Text Wrap") do |s|
              s.set_print_option(:grid_lines, true)
              s.set_column(0, width: 15)
              s.add_row(["This is a long sentence that wraps inside the cell."], styles: { 0 => "wrap" })
            end
          end
        RUBY

        "align_indent.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_indent.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_style("indent_1") { |s| s.indent(1) }
            w.add_style("indent_3") { |s| s.indent(3) }
            w.add_sheet("Indent") do |s|
              s.set_print_option(:grid_lines, true)
              s.add_row(["No Indent"])
              s.add_row(["Indent 1"], styles: { 0 => "indent_1" })
              s.add_row(["Indent 3"], styles: { 0 => "indent_3" })
            end
          end
        RUBY

        "col_widths.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "col_widths.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Widths") do |s|
              s.set_column(0, width: 30)
              s.set_column(1, width: 10)
              s.add_row(["Wide Column A", "Narrow B"])
            end
          end
        RUBY

        "row_heights.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "row_heights.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Heights") do |s|
              s.add_row(["Normal Row"])
              s.add_row(["Tall Row"], height: 40)
            end
          end
        RUBY

        "row_grouping.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "row_grouping.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Row Grouping") do |s|
              s.add_row(["Parent Row 1"])
              s.add_row(["Child Row 1.1"], outline_level: 1)
              s.add_row(["Child Row 1.2"], outline_level: 1)
            end
          end
        RUBY

        "col_grouping.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "col_grouping.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Col Grouping") do |s|
              s.set_column(0, outline_level: 0)
              s.set_column(1, outline_level: 1)
              s.set_column(2, outline_level: 1)
              s.add_row(["Col A", "Col B (Grouped)", "Col C (Grouped)"])
            end
          end
        RUBY

        "chart_line.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_line.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data") do |s|
              s.add_row(["Day", "Value"])
              s.add_row(["Mon", 10])
              s.add_row(["Tue", 15])
              s.add_row(["Wed", 12])
              s.add_chart(
                type: :line,
                title: "Daily Value",
                from_col: 3, from_row: 0, to_col: 8, to_row: 12,
                series: [{ cat_ref: "Data!$A$2:$A$4", val_ref: "Data!$B$2:$B$4" }]
              )
            end
          end
        RUBY

        "chart_area.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_area.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data") do |s|
              s.add_row(["Day", "Value"])
              s.add_row(["Mon", 10])
              s.add_row(["Tue", 15])
              s.add_chart(
                type: :area,
                title: "Daily Area",
                from_col: 3, from_row: 0, to_col: 8, to_row: 12,
                series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }]
              )
            end
          end
        RUBY

        "chart_pie.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_pie.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data") do |s|
              s.add_row(["Label", "Percent"])
              s.add_row(["Yes", 70])
              s.add_row(["No", 30])
              s.add_chart(
                type: :pie,
                title: "Responses",
                from_col: 3, from_row: 0, to_col: 8, to_row: 12,
                series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }]
              )
            end
          end
        RUBY

        "chart_doughnut.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_doughnut.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data") do |s|
              s.add_row(["Label", "Percent"])
              s.add_row(["A", 40])
              s.add_row(["B", 60])
              s.add_chart(
                type: :doughnut,
                title: "Ratio",
                from_col: 3, from_row: 0, to_col: 8, to_row: 12,
                series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }]
              )
            end
          end
        RUBY

        "chart_scatter.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_scatter.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data") do |s|
              s.add_row(["X", "Y"])
              s.add_row([1, 10])
              s.add_row([2, 15])
              s.add_chart(
                type: :scatter,
                title: "Scatter Plot",
                from_col: 3, from_row: 0, to_col: 8, to_row: 12,
                series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }]
              )
            end
          end
        RUBY

        "chart_radar.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_radar.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data") do |s|
              s.add_row(["Stat", "Value"])
              s.add_row(["Atk", 80])
              s.add_row(["Def", 60])
              s.add_chart(
                type: :radar,
                title: "Stats",
                from_col: 3, from_row: 0, to_col: 8, to_row: 12,
                series: [{ cat_ref: "Data!$A$2:$A$3", val_ref: "Data!$B$2:$B$3" }]
              )
            end
          end
        RUBY

        "sparkline_line.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "sparkline_line.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Sparkline") do |s|
              s.add_row([10, 20, 15, 30, nil])
              s.add_sparkline_group(
                type: :line,
                sparklines: [{ location_ref: "E1", data_ref: "Sparkline!A1:D1" }]
              )
            end
          end
        RUBY

        "sparkline_column.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "sparkline_column.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Sparkline") do |s|
              s.add_row([5, 12, 8, 15, nil])
              s.add_sparkline_group(
                type: :column,
                sparklines: [{ location_ref: "E1", data_ref: "Sparkline!A1:D1" }]
              )
            end
          end
        RUBY

        "cf_color_scale.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cf_color_scale.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Colors") do |s|
              s.add_row([10])
              s.add_row([50])
              s.add_row([90])
              s.add_conditional_format("A1:A3", type: :colorScale, priority: 1)
            end
          end
        RUBY

        "cf_data_bar.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cf_data_bar.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Data Bars") do |s|
              s.add_row([20])
              s.add_row([60])
              s.add_row([100])
              s.add_conditional_format("A1:A3", type: :dataBar, priority: 1, color: "FF0070C0")
            end
          end
        RUBY

        "cf_icon_set.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cf_icon_set.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Icons") do |s|
              s.add_row([25])
              s.add_row([50])
              s.add_row([75])
              s.add_conditional_format("A1:A3", type: :iconSet, icon_style: "3Arrows", priority: 1)
            end
          end
        RUBY

        "interactive_autofilter.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_autofilter.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Filter") do |s|
              s.add_row(["Name", "Department"])
              s.add_row(["Alice", "HR"])
              s.add_row(["Bob", "Eng"])
              s.set_auto_filter("A1:B3")
            end
            w.add_defined_name("_xlnm._FilterDatabase", "Filter!$A$1:$B$3", sheet: "Filter", hidden: true)
          end
        RUBY

        "interactive_validation_list.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_validation_list.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("List Validation") do |s|
              s.add_row(["Department", "Select:"])
              s.add_data_validation("B2", type: "list", formula1: '"HR,Sales,Engineering"')
            end
          end
        RUBY

        "interactive_validation_range.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_validation_range.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Range Validation") do |s|
              s.add_row(["Age", "Enter (18-99):"])
              s.add_data_validation("B2", type: "whole", operator: "between", formula1: "18", formula2: "99", show_error_message: true, error_title: "Invalid Age", error: "Age must be between 18 and 99!")
            end
          end
        RUBY

        "interactive_comments.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_comments.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Comments") do |s|
              s.add_row(["Item A", "Item B"])
              s.add_comment("A1", "This is an important comment.", author: "System")
            end
          end
        RUBY

        "embedded_images.rb" => <<~RUBY
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "embedded_images.xlsx"

          # Generate a tiny 10x10 dummy red PNG file in memory using pure ruby
          # (PNG signature + IHDR chunk + IDAT chunk + IEND chunk)
          dummy_png = [
            0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a,
            0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x0a, 0x00, 0x00, 0x00, 0x0a,
            0x08, 0x02, 0x00, 0x00, 0x00, 0x02, 0xeb, 0x8a,
            0x07, 0x00, 0x00, 0x00, 0x14, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9c, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
            0x06, 0x12, 0x03, 0x00, 0x12, 0x00, 0x01, 0xdc,
            0x0b, 0x7e, 0x22, 0x1e, 0x00, 0x00, 0x00, 0x00,
            0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82
          ].pack("C*")

          Xlsxrb.generate(output_path) do |w|
            w.add_sheet("Images") do |s|
              s.add_row(["Logo Target cell:"])
              s.add_image(dummy_png, ext: "png", from_col: 1, from_row: 1, to_col: 3, to_row: 5)
            end
          end
        RUBY
      }.freeze

      def self.write_all
        target_dir = File.expand_path("../../../examples/visual", __dir__)
        FileUtils.mkdir_p(target_dir)

        FILES.each do |filename, code|
          path = File.join(target_dir, filename)
          File.write(path, code)
          puts "Generated visual example script: #{filename}"
        end
      end
    end
  end
end
