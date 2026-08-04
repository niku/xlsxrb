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
            w.style("currency") { |s| s.num_fmt("$#,##0.00") }
            w.style("percent") { |s| s.num_fmt("0.0%") }
            w.sheet("Numbers") do |s|
              s.row(["Format", "Value"])
              s.row(["Integer", 12345])
              s.row(["Float", 123.456])
              s.row(["Currency", 1234.5], styles: { 1 => "currency" })
              s.row(["Percentage", 0.85], styles: { 1 => "percent" })
            end
          end
        RUBY

        "cell_dates.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          require "date"
          output_path = ARGV[0] || "cell_dates.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("custom_date") { |s| s.num_fmt("yyyy-mm-dd") }
            w.sheet("Dates") do |s|
              s.row(["Format", "Date Value"])
              s.row(["Default Date", Date.new(2026, 7, 1)])
              s.row(["Formatted Date", Date.new(2026, 12, 25)], styles: { 1 => "custom_date" })
            end
          end
        RUBY

        "cell_times.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          require "time"
          output_path = ARGV[0] || "cell_times.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("time_fmt") { |s| s.num_fmt("hh:mm:ss") }
            w.sheet("Times") do |s|
              s.row(["Format", "Time Value"])
              s.row(["DateTime", Time.new(2026, 7, 1, 12, 34, 56)])
              s.row(["Time Only", Time.new(2026, 7, 1, 9, 15, 0)], styles: { 1 => "time_fmt" })
            end
          end
        RUBY

        "cell_booleans.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cell_booleans.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Booleans") do |s|
              s.row(["Label", "Boolean Value"])
              s.row(["Is Active", true])
              s.row(["Is Pending", false])
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
            w.sheet("Rich Text") do |s|
              s.row(["Format", "Value"])
              s.row(["Rich Text Cell", rt])
            end
          end
        RUBY

        "cell_formulas.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cell_formulas.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Formulas") do |s|
              s.row(["Item", "Value"])
              s.row(["A", 10])
              s.row(["B", 20])
              s.row(["SUM", "=SUM(B2:B3)"])
              s.row(["AVERAGE", "=AVERAGE(B2:B3)"])
            end
          end
        RUBY

        "font_families.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_families.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("arial") { |s| s.font_name("Arial") }
            w.style("times") { |s| s.font_name("Times New Roman") }
            w.style("courier") { |s| s.font_name("Courier New") }
            w.sheet("Fonts") do |s|
              s.row(["Font Family", "Preview"])
              s.row(["Arial", "Hello Arial"], styles: { 1 => "arial" })
              s.row(["Times New Roman", "Hello Times"], styles: { 1 => "times" })
              s.row(["Courier New", "Hello Courier"], styles: { 1 => "courier" })
            end
          end
        RUBY

        "font_sizes.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_sizes.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("size_10") { |s| s.size(10) }
            w.style("size_16") { |s| s.size(16) }
            w.style("size_24") { |s| s.size(24) }
            w.sheet("Font Sizes") do |s|
              s.row(["Size", "Text"])
              s.row(["10pt", "Small Text"], styles: { 1 => "size_10" })
              s.row(["16pt", "Medium Text"], styles: { 1 => "size_16" })
              s.row(["24pt", "Large Text"], styles: { 1 => "size_24" })
            end
          end
        RUBY

        "font_colors.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_colors.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("red") { |s| s.font_color("FFC00000") }
            w.style("green") { |s| s.font_color("FF00B050") }
            w.style("blue") { |s| s.font_color("FF0070C0") }
            w.sheet("Colors") do |s|
              s.row(["Color", "Preview"])
              s.row(["Red", "Red Text"], styles: { 1 => "red" })
              s.row(["Green", "Green Text"], styles: { 1 => "green" })
              s.row(["Blue", "Blue Text"], styles: { 1 => "blue" })
            end
          end
        RUBY

        "font_styles.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_styles.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("bold") { |s| s.bold }
            w.style("italic") { |s| s.italic }
            w.style("underline") { |s| s.underline }
            w.style("strike") { |s| s.strike }
            w.sheet("Styles") do |s|
              s.row(["Style", "Preview"])
              s.row(["Bold", "Bold Text"], styles: { 1 => "bold" })
              s.row(["Italic", "Italic Text"], styles: { 1 => "italic" })
              s.row(["Underline", "Underlined Text"], styles: { 1 => "underline" })
              s.row(["Strike-through", "Struck Text"], styles: { 1 => "strike" })
            end
          end
        RUBY

        "font_vertical_align.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "font_vertical_align.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("super") { |s| s.vert_align("superscript") }
            w.style("sub") { |s| s.vert_align("subscript") }
            w.sheet("Vertical Align") do |s|
              s.row(["Format", "Text"])
              s.row(["Superscript", "x2"], styles: { 1 => "super" })
              s.row(["Subscript", "H2O"], styles: { 1 => "sub" })
            end
          end
        RUBY

        "border_thin.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_thin.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("thin") { |s| s.border_left(style: "thin").border_right(style: "thin").border_top(style: "thin").border_bottom(style: "thin") }
            w.sheet("Thin Borders") do |s|
              s.row(["Normal Cell", "Thin Border Cell"], styles: { 1 => "thin" })
            end
          end
        RUBY

        "border_medium.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_medium.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("medium") { |s| s.border_left(style: "medium").border_right(style: "medium").border_top(style: "medium").border_bottom(style: "medium") }
            w.sheet("Medium Borders") do |s|
              s.row(["Normal Cell", "Medium Border Cell"], styles: { 1 => "medium" })
            end
          end
        RUBY

        "border_dashed.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_dashed.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("dashed") { |s| s.border_left(style: "dashed").border_right(style: "dashed").border_top(style: "dashed").border_bottom(style: "dashed") }
            w.sheet("Dashed Borders") do |s|
              s.row(["Normal Cell", "Dashed Border Cell"], styles: { 1 => "dashed" })
            end
          end
        RUBY

        "border_double.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_double.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("double") { |s| s.border_left(style: "double").border_right(style: "double").border_top(style: "double").border_bottom(style: "double") }
            w.sheet("Double Borders") do |s|
              s.row(["Normal Cell", "Double Border Cell"], styles: { 1 => "double" })
            end
          end
        RUBY

        "border_slanted.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "border_slanted.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("slanted") { |s| s.border_left(style: "slantedDashDot").border_right(style: "slantedDashDot").border_top(style: "slantedDashDot").border_bottom(style: "slantedDashDot") }
            w.sheet("Slanted Borders") do |s|
              s.row(["Normal Cell", "Slanted Border Cell"], styles: { 1 => "slanted" })
            end
          end
        RUBY

        "fill_solid_colors.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "fill_solid_colors.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("red_fill") { |s| s.fill_color("FFFFC7CE") }
            w.style("green_fill") { |s| s.fill_color("FFC6EFCE") }
            w.sheet("Fills") do |s|
              s.row(["Color", "Preview"])
              s.row(["Red", "Red Fill"], styles: { 1 => "red_fill" })
              s.row(["Green", "Green Fill"], styles: { 1 => "green_fill" })
            end
          end
        RUBY

        "fill_patterns.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "fill_patterns.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("dark_gray") { |s| s.fill(pattern: "darkGray", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
            w.style("grid_fill") { |s| s.fill(pattern: "darkGrid", fg_color: "FFC0C0C0", bg_color: "FFFFFFFF") }
            w.sheet("Patterns") do |s|
              s.row(["Pattern", "Preview"])
              s.row(["Dark Gray", "Pattern Fill"], styles: { 1 => "dark_gray" })
              s.row(["Dark Grid", "Grid Fill"], styles: { 1 => "grid_fill" })
            end
          end
        RUBY

        "fill_gradients.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "fill_gradients.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("gradient") do |style|
              style.fill_gradient(type: "linear", degree: 45, stops: [
                { position: 0, color: "FFFFFFFF" },
                { position: 1, color: "FF4F81BD" }
              ])
            end
            w.sheet("Gradients") do |s|
              s.row(["Normal Cell", "Gradient Cell"], styles: { 1 => "gradient" })
            end
          end
        RUBY

        "align_horizontal.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_horizontal.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("left") { |s| s.align_horizontal("left") }
            w.style("center") { |s| s.align_horizontal("center") }
            w.style("right") { |s| s.align_horizontal("right") }
            w.sheet("Alignment") do |s|
              s.set_print_option(:grid_lines, true)
              s.column(0, width: 20)
              s.column(1, width: 20)
              s.column(2, width: 20)
              s.row(["Left", "Center", "Right"], styles: { 0 => "left", 1 => "center", 2 => "right" })
            end
          end
        RUBY

        "align_vertical.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_vertical.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("top") { |s| s.align_vertical("top") }
            w.style("center") { |s| s.align_vertical("center") }
            w.style("bottom") { |s| s.align_vertical("bottom") }
            w.sheet("Vertical Alignment") do |s|
              s.set_print_option(:grid_lines, true)
              s.row(["Top", "Center", "Bottom"], styles: { 0 => "top", 1 => "center", 2 => "bottom" }, height: 40)
            end
          end
        RUBY

        "align_text_rotation.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_text_rotation.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("rot_45") { |s| s.text_rotation(45) }
            w.style("rot_90") { |s| s.text_rotation(90) }
            w.sheet("Rotation") do |s|
              s.set_print_option(:grid_lines, true)
              s.row(["Rotated 45", "Rotated 90"], styles: { 0 => "rot_45", 1 => "rot_90" }, height: 50)
            end
          end
        RUBY

        "align_text_wrap.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_text_wrap.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("wrap") { |s| s.wrap_text }
            w.sheet("Text Wrap") do |s|
              s.set_print_option(:grid_lines, true)
              s.column(0, width: 15)
              s.row(["This is a long sentence that wraps inside the cell."], styles: { 0 => "wrap" })
            end
          end
        RUBY

        "align_indent.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "align_indent.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.style("indent_1") { |s| s.indent(1) }
            w.style("indent_3") { |s| s.indent(3) }
            w.sheet("Indent") do |s|
              s.set_print_option(:grid_lines, true)
              s.row(["No Indent"])
              s.row(["Indent 1"], styles: { 0 => "indent_1" })
              s.row(["Indent 3"], styles: { 0 => "indent_3" })
            end
          end
        RUBY

        "col_widths.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "col_widths.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Widths") do |s|
              s.column(0, width: 30)
              s.column(1, width: 10)
              s.row(["Wide Column A", "Narrow B"])
            end
          end
        RUBY

        "row_heights.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "row_heights.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Heights") do |s|
              s.row(["Normal Row"])
              s.row(["Tall Row"], height: 40)
            end
          end
        RUBY

        "row_grouping.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "row_grouping.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Row Grouping") do |s|
              s.row(["Parent Row 1"])
              s.row(["Child Row 1.1"], outline_level: 1)
              s.row(["Child Row 1.2"], outline_level: 1)
            end
          end
        RUBY

        "col_grouping.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "col_grouping.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Col Grouping") do |s|
              s.column(0, outline_level: 0)
              s.column(1, outline_level: 1)
              s.column(2, outline_level: 1)
              s.row(["Col A", "Col B (Grouped)", "Col C (Grouped)"])
            end
          end
        RUBY

        "chart_line.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "chart_line.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Data") do |s|
              s.row(["Day", "Value"])
              s.row(["Mon", 10])
              s.row(["Tue", 15])
              s.row(["Wed", 12])
              s.chart(
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
            w.sheet("Data") do |s|
              s.row(["Day", "Value"])
              s.row(["Mon", 10])
              s.row(["Tue", 15])
              s.chart(
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
            w.sheet("Data") do |s|
              s.row(["Label", "Percent"])
              s.row(["Yes", 70])
              s.row(["No", 30])
              s.chart(
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
            w.sheet("Data") do |s|
              s.row(["Label", "Percent"])
              s.row(["A", 40])
              s.row(["B", 60])
              s.chart(
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
            w.sheet("Data") do |s|
              s.row(["X", "Y"])
              s.row([1, 10])
              s.row([2, 15])
              s.chart(
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
            w.sheet("Data") do |s|
              s.row(["Stat", "Value"])
              s.row(["Atk", 80])
              s.row(["Def", 60])
              s.chart(
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
            w.sheet("Sparkline") do |s|
              s.row([10, 20, 15, 30, nil])
              s.sparkline_group(
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
            w.sheet("Sparkline") do |s|
              s.row([5, 12, 8, 15, nil])
              s.sparkline_group(
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
            w.sheet("Colors") do |s|
              s.row([10])
              s.row([50])
              s.row([90])
              s.conditional_format("A1:A3", type: :colorScale, priority: 1)
            end
          end
        RUBY

        "cf_data_bar.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cf_data_bar.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Data Bars") do |s|
              s.row([20])
              s.row([60])
              s.row([100])
              s.conditional_format("A1:A3", type: :dataBar, priority: 1, color: "FF0070C0")
            end
          end
        RUBY

        "cf_icon_set.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "cf_icon_set.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Icons") do |s|
              s.row([25])
              s.row([50])
              s.row([75])
              s.conditional_format("A1:A3", type: :iconSet, icon_style: "3Arrows", priority: 1)
            end
          end
        RUBY

        "interactive_autofilter.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_autofilter.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Filter") do |s|
              s.row(["Name", "Department"])
              s.row(["Alice", "HR"])
              s.row(["Bob", "Eng"])
              s.auto_filter("A1:B3")
            end
            w.add_defined_name("_xlnm._FilterDatabase", "Filter!$A$1:$B$3", sheet: "Filter", hidden: true)
          end
        RUBY

        "interactive_validation_list.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_validation_list.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("List Validation") do |s|
              s.row(["Department", "Select:"])
              s.validate_data("B2", type: "list", formula1: '"HR,Sales,Engineering"')
            end
          end
        RUBY

        "interactive_validation_range.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_validation_range.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Range Validation") do |s|
              s.row(["Age", "Enter (18-99):"])
              s.validate_data("B2", type: "whole", operator: "between", formula1: "18", formula2: "99", show_error_message: true, error_title: "Invalid Age", error: "Age must be between 18 and 99!")
            end
          end
        RUBY

        "interactive_comments.rb" => <<~RUBY,
          # frozen_string_literal: true
          require "xlsxrb"
          output_path = ARGV[0] || "interactive_comments.xlsx"
          Xlsxrb.generate(output_path) do |w|
            w.sheet("Comments") do |s|
              s.row(["Item A", "Item B"])
              s.comment("A1", "This is an important comment.", author: "System")
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
            w.sheet("Images") do |s|
              s.row(["Logo Target cell:"])
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
