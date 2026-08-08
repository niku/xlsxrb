# frozen_string_literal: true

require "fileutils"
require "open3"

module Xlsxrb
  module Visual
    class GalleryGenerator
      EXPLANATIONS = {
        "basic_data" => "Demonstrates simple tabular data writing with basic Ruby types (Strings, Numbers, Dates, Booleans).",
        "styles_fonts_fills" => "Demonstrates cell formatting, including custom font sizing, bold/italic text, custom text colors, and background fill colors.",
        "borders" => "Demonstrates border styles (thin, medium, thick, hair, dashed, medium dashed, dotted, double, dash-dot, medium dash-dot, dash-dot-dot, slanted, and diagonal cross borders) applied to cell ranges.",
        "fonts" => "Demonstrates cell fonts properties (Arial, Georgia, Courier New, Times New Roman, Tahoma, sizes 10pt/16pt/24pt, red/green/blue colors, bold/italic/underline/double underline/strike-through styles, superscript/subscript).",
        "merge_freeze" => "Demonstrates merging a cell range into a single cell, and freezing the top rows of a sheet.",
        "conditional_formatting" => "Demonstrates adding conditional formatting rules that style cells automatically based on value ranges.",
        "japanese_text" => "Demonstrates writing multi-byte Japanese text and setting appropriate Japanese font names (e.g., Noto Sans CJK JP).",
        "chart_bar" => "Demonstrates embedding a standard 2D Bar Chart referencing worksheet cell ranges.",
        "cell_numbers" => "Demonstrates custom formatting for integers, floating point numbers, currencies, and percentages.",
        "cell_dates" => "Demonstrates dates serialized natively and formatted with standard or custom format strings.",
        "cell_times" => "Demonstrates timestamp values serialized natively and formatted showing hours, minutes, and seconds.",
        "cell_booleans" => "Demonstrates boolean values serialized and rendered.",
        "cell_rich_text" => "Demonstrates Rich Text cells with multiple font weights, styles, and colors in a single cell.",
        "cell_formulas" => "Demonstrates standard spreadsheet calculations and formulas (SUM, AVERAGE).",
        "fill_solid_colors" => "Demonstrates solid cell background fill colors.",
        "fill_patterns" => "Demonstrates standard pattern fills (darkGray, darkGrid) in cell backgrounds.",
        "fill_gradients" => "Demonstrates linear gradients inside cell backgrounds.",
        "align_horizontal" => "Demonstrates horizontal text alignment (left, center, right).",
        "align_vertical" => "Demonstrates vertical text alignment (top, center, bottom).",
        "align_text_rotation" => "Demonstrates text rotated by specific angles (45, 90 degrees).",
        "align_text_wrap" => "Demonstrates auto-wrapping multi-line text inside narrow cells.",
        "align_indent" => "Demonstrates text indentation inside cells.",
        "col_widths" => "Demonstrates setting custom column widths.",
        "row_heights" => "Demonstrates setting custom row heights.",
        "row_grouping" => "Demonstrates outline grouping for rows.",
        "col_grouping" => "Demonstrates outline grouping for columns.",
        "chart_line" => "Demonstrates embedding a standard 2D Line Chart.",
        "chart_area" => "Demonstrates embedding a standard 2D Area Chart.",
        "chart_pie" => "Demonstrates embedding a standard 2D Pie Chart.",
        "chart_doughnut" => "Demonstrates embedding a standard 2D Doughnut Chart.",
        "chart_scatter" => "Demonstrates embedding a standard 2D Scatter Plot.",
        "chart_radar" => "Demonstrates embedding a standard 2D Radar Chart.",
        "sparkline_line" => "Demonstrates embedded line sparklines in cell ranges.",
        "sparkline_column" => "Demonstrates embedded column sparklines in cell ranges.",
        "cf_color_scale" => "Demonstrates color scale/heatmap conditional formatting.",
        "cf_data_bar" => "Demonstrates data bar visual conditional formatting indicators.",
        "cf_icon_set" => "Demonstrates icon set indicators (red/yellow/green arrows).",
        "interactive_autofilter" => "Demonstrates enabling auto-filter sorting headers on tables. Download the sheet to interactively filter and sort columns.",
        "interactive_validation_list" => "Demonstrates dropdown list data validations. Open the sheet in Excel and select cell A1/A2 to see the dropdown list in action.",
        "interactive_validation_range" => "Demonstrates range constraints for whole number validations. Open the sheet in Excel and try entering a value outside 10-100 to trigger the warning.",
        "interactive_comments" => "Demonstrates cell pop-up comments. Open the sheet in Excel and hover your mouse over the cell with the red triangle to view the comment.",
        "embedded_images" => "Demonstrates embedding raster PNG images in cell ranges.",
        "cell_num_scientific" => "Demonstrates scientific number formats (0.00E+00).",
        "cell_num_fractions" => "Demonstrates fraction number formats (# ?/?).",
        "cell_num_percent_decimals" => "Demonstrates percentages with two decimal places (0.00%).",
        "cell_num_custom_colors" => "Demonstrates custom colored formats for positive and negative numbers.",
        "align_horizontal_fill" => "Demonstrates horizontal fill alignment (repeats value to fill cell width).",
        "align_horizontal_justify" => "Demonstrates horizontal justify text alignment.",
        "col_width_tall" => "Demonstrates setting very wide column widths.",
        "row_height_tall" => "Demonstrates setting very tall row heights.",
        "sheet_tab_colors" => "Demonstrates customizing tab colors of individual worksheets.",
        "workbook_three_sheets" => "Demonstrates creating workbooks with multiple worksheets.",
        "chart_bar_stacked" => "Demonstrates embedding a stacked 2D Bar Chart.",
        "chart_bar_percent_stacked" => "Demonstrates embedding a 100% stacked 2D Bar Chart.",
        "chart_line_stacked" => "Demonstrates embedding a stacked 2D Line Chart.",
        "chart_area_stacked" => "Demonstrates embedding a stacked 2D Area Chart.",
        "cf_cell_greater_than" => "Demonstrates conditional formatting highlighting cells greater than a threshold.",
        "cf_cell_less_than" => "Demonstrates conditional formatting highlighting cells less than a threshold.",
        "cf_cell_equal_to" => "Demonstrates conditional formatting highlighting cells equal to a target value.",
        "cf_cell_between" => "Demonstrates conditional formatting highlighting cells within a range.",
        "cf_contains_text" => "Demonstrates conditional formatting highlighting cells containing specific text.",
        "cf_begins_with" => "Demonstrates conditional formatting highlighting cells starting with specific text.",
        "cf_ends_with" => "Demonstrates conditional formatting highlighting cells ending with specific text.",
        "cf_cell_greater_equal" => "Demonstrates conditional formatting highlighting cells greater than or equal to a threshold.",
        "cf_expression_formula" => "Demonstrates conditional formatting using a custom formula expression.",
        "page_orientation_landscape" => "Demonstrates landscape page setup for printing.",
        "page_paper_size_a3" => "Demonstrates setting paper size to A3 (paper size 8).",
        "page_margins_wide" => "Demonstrates setting wide page margins.",
        "page_margins_narrow" => "Demonstrates setting narrow page margins.",
        "page_header_footer" => "Demonstrates setting odd page headers and footers.",
        "page_grid_lines_print" => "Demonstrates enabling printing of grid lines.",
        "page_headings_print" => "Demonstrates enabling printing of row and column headings.",
        "view_show_grid_lines" => "Demonstrates disabling visible grid lines in spreadsheet view.",
        "view_zoom_scale" => "Demonstrates setting custom zoom scale in sheet view (e.g. 150%).",
        "interactive_validation_date" => "Demonstrates interactive date range constraints validation rules.",
        "interactive_validation_text_length" => "Demonstrates interactive text length validation rules.",
        "interactive_validation_custom" => "Demonstrates interactive custom formula validation rules.",
        "interactive_validation_time" => "Demonstrates interactive time validation rules.",
        "cell_num_currency_jpy" => "Demonstrates Yen Currency format code formatting."
      }.freeze

      def self.generate
        require_relative "screenshot_capturer"

        gallery_dir = File.expand_path("../../../docs/visual", __dir__)
        images_dir = File.join(gallery_dir, "images")
        files_dir = File.join(gallery_dir, "files")
        baselines_dir = File.expand_path("../../../test/visual/baselines", __dir__)
        illustrations_dir = File.expand_path("illustrations", __dir__)

        # 0. Clean and recreate target directories
        FileUtils.rm_rf(images_dir)
        FileUtils.rm_rf(files_dir)
        FileUtils.mkdir_p(images_dir)
        FileUtils.mkdir_p(files_dir)

        # 1. Run all examples to generate XLSX files for download
        examples = Dir.glob(File.expand_path("../../../examples/visual/*.rb", __dir__))
        examples.each do |example_path|
          name = File.basename(example_path, ".rb")
          xlsx_path = File.join(files_dir, "#{name}.xlsx")
          system("ruby", "-Ilib", example_path, xlsx_path)
        end

        # 2. Capture interactive GUI screenshots if tools are available
        ScreenshotCapturer.capture_all

        # 3. Build gallery items by copying images from baselines
        items = []

        examples.each do |example_path|
          name = File.basename(example_path, ".rb")
          title = name.split("_").map(&:capitalize).join(" ")
          explanation = EXPLANATIONS[name] || "Visual demonstration for #{title}."

          puts "Generating gallery item: #{title}..."

          # Copy baseline PNGs to docs/visual/images/
          baseline_dir = File.join(baselines_dir, name)
          copied_pngs = []

          if File.directory?(baseline_dir)
            baseline_pngs = Dir.glob(File.join(baseline_dir, "page-*.png")).sort_by do |path|
              path.match(/page-(\d+)\.png/)[1].to_i
            end

            baseline_pngs.each_with_index do |png_path, idx|
              dest_name = "#{name}_page-#{idx + 1}.png"
              dest_path = File.join(images_dir, dest_name)
              FileUtils.cp(png_path, dest_path)
              copied_pngs << "../../test/visual/baselines/#{name}/page-#{idx + 1}.png"
            end
          else
            puts "  WARNING: No baseline found for #{name}, skipping images."
          end

          # If an illustration exists (e.g. interactive UI screenshots), copy it as an additional page
          illustration_src = File.join(illustrations_dir, "#{name}_page-2.png")
          if File.exist?(illustration_src)
            dest_name = "#{name}_page-#{copied_pngs.size + 1}.png"
            dest_path = File.join(images_dir, dest_name)
            FileUtils.cp(illustration_src, dest_path)
            copied_pngs << "../../test/visual/support/illustrations/#{name}_page-2.png"
          end

          # Read example code
          code = File.read(example_path)

          # Run example to capture console output
          xlsx_path = File.join(files_dir, "#{name}.xlsx")
          stdout, _stderr, _status = Open3.capture3("ruby", "-Ilib", example_path, xlsx_path)
          console_output = stdout && !stdout.strip.empty? ? stdout : nil

          items << {
            name: name,
            title: title,
            explanation: explanation,
            code: code,
            pngs: copied_pngs,
            console_output: console_output
          }
        end

        # Build Markdown
        markdown = +""
        markdown << "# Visual Examples Gallery\n\n"
        markdown << "This gallery showcases `xlsxrb` DSL usage side-by-side with the visual rendering in LibreOffice Calc.\n\n"
        markdown << "[◄ Back to README](../../README.md)\n\n"

        # 1. Add Visual Capability Grid (TOC)
        markdown << "## Capability Overview\n\n"
        markdown << "<table>\n"
        markdown << "<thead>\n"
        markdown << "<tr>\n"
        markdown << "<th align=\"left\">Feature</th>\n"
        markdown << "<th align=\"center\">Visual Preview</th>\n"
        markdown << "<th align=\"center\">Link to Detail</th>\n"
        markdown << "</tr>\n"
        markdown << "</thead>\n"
        markdown << "<tbody>\n"
        items.each do |item|
          first_page = item[:pngs].first
          anchor = item[:title].downcase.gsub(" ", "-")
          markdown << "<tr>\n"
          markdown << "<td><strong>#{item[:title]}</strong></td>\n"
          markdown << "<td align=\"center\"><img src=\"#{first_page}\" width=\"160\" alt=\"#{item[:title]}\"/></td>\n"
          markdown << "<td align=\"center\"><a href=\"##{anchor}\">View Code &amp; Detail</a></td>\n"
          markdown << "</tr>\n"
        end
        markdown << "</tbody>\n"
        markdown << "</table>\n\n"

        markdown << "\n---\n\n"

        # 2. Add Detailed Sections
        items.each do |item|
          markdown << "## #{item[:title]}\n\n"

          if item[:name].start_with?("interactive_")
            markdown << "> [!TIP]\n"
            markdown << "> **Interactive Feature**: This example uses interactive Excel behaviors (such as validation dropdowns, comments, or autofilters). Since the static visual preview below represents a printed page layout, please use the **Live Preview** to interact with it!\n\n"
          end

          markdown << "#{item[:explanation]}\n\n"

          markdown << "### Rendered Output (LibreOffice Calc)\n\n"
          item[:pngs].each do |png_path|
            markdown << "<div><img src=\"#{png_path}\" width=\"100%\" alt=\"Preview\"/></div>\n\n"
          end

          markdown << "### DSL Code\n\n"
          markdown << "```ruby\n"
          markdown << item[:code].strip << "\n"
          markdown << "```\n\n"

          if item[:console_output]
            markdown << "### Console Output\n\n"
            markdown << "```text\n"
            markdown << item[:console_output].strip << "\n"
            markdown << "```\n\n"
          end

          markdown << "<hr/>\n\n"
        end

        # Write VisualGallery.md
        File.write(File.join(gallery_dir, "VisualGallery.md"), markdown)
        puts "Gallery generation complete: docs/visual/VisualGallery.md"
      end
    end
  end
end
