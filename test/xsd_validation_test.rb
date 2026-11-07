# frozen_string_literal: true

require "test_helper"
require "nokogiri"
require "zip"
require "tempfile"

class XsdValidationTest < Test::Unit::TestCase
  SCHEMAS = Dir.chdir(File.expand_path("fixtures/xsd", __dir__)) do
    {
      sml: Nokogiri::XML::Schema(File.read("sml.xsd")),
      drawing: Nokogiri::XML::Schema(File.read("dml-spreadsheetDrawing.xsd")),
      chart: Nokogiri::XML::Schema(File.read("dml-chart.xsd")),
      app_props: Nokogiri::XML::Schema(File.read("shared-documentPropertiesExtended.xsd")),
      custom_props: Nokogiri::XML::Schema(File.read("shared-documentPropertiesCustom.xsd"))
    }
  end

  # Helper to validate all relevant XML parts in an XLSX file against ECMA-376 XSD schemas
  def validate_xlsx_file(file_path)
    Zip::File.open(file_path) do |zip|
      zip.each do |entry|
        next unless entry.file?
        next unless entry.name.end_with?(".xml")

        xml_content = entry.get_input_stream.read
        doc = Nokogiri::XML(xml_content)

        case entry.name
        when %r{\Axl/workbook\.xml\z},
             %r{\Axl/worksheets/sheet\d+\.xml\z},
             %r{\Axl/styles\.xml\z},
             %r{\Axl/sharedStrings\.xml\z},
             %r{\Axl/tables/table\d+\.xml\z},
             %r{\Axl/comments\d+\.xml\z},
             %r{\Axl/pivotTables/pivotTable\d+\.xml\z},
             %r{\Axl/pivotCache/pivotCacheDefinition\d+\.xml\z}
          errors = SCHEMAS[:sml].validate(doc)
          assert_empty errors, "#{entry.name} failed SML XSD validation: #{errors.join(", ")}"

        when %r{\Axl/drawings/drawing\d+\.xml\z}
          errors = SCHEMAS[:drawing].validate(doc)
          assert_empty errors, "#{entry.name} failed DrawingML XSD validation: #{errors.join(", ")}"

        when %r{\Axl/charts/chart\d+\.xml\z}
          errors = SCHEMAS[:chart].validate(doc)
          assert_empty errors, "#{entry.name} failed ChartML XSD validation: #{errors.join(", ")}"

        when %r{\AdocProps/app\.xml\z}
          errors = SCHEMAS[:app_props].validate(doc)
          assert_empty errors, "#{entry.name} failed App Properties XSD validation: #{errors.join(", ")}"

        when %r{\AdocProps/custom\.xml\z}
          errors = SCHEMAS[:custom_props].validate(doc)
          assert_empty errors, "#{entry.name} failed Custom Properties XSD validation: #{errors.join(", ")}"
        end
      end
    end
  end

  # 1. Basic workbook and multiple sheets
  def test_basic_workbook_and_multiple_sheets
    tmp = Tempfile.new(["xsd_basic", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(["Hello", 123, 45.67, true, false, nil])
        end
        w.sheet("Sheet2") do |s|
          s.row(["Second Sheet", Date.today])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 2. Rich styles (fonts, fills, borders, alignments, numfmts)
  def test_styles_xml_validation
    tmp = Tempfile.new(["xsd_styles", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.style("header",
                font: { bold: true, size: 14, color: "FFFFFFFF", font_name: "Arial" },
                fill: { pattern: "solid", fg_color: "FF4F81BD" },
                alignment: { horizontal: "center", vertical: "center", wrap_text: true },
                border: { bottom: { style: "double", color: "FF000000" } })

        w.style("currency_cell",
                num_fmt: "$#,##0.00;($#,##0.00);\"-\"",
                font: { italic: true },
                border: {
                  left: { style: "thin", color: "FFCCCCCC" },
                  right: { style: "thin", color: "FFCCCCCC" }
                })

        w.sheet("StylesSheet") do |s|
          s.row(["Header 1", "Header 2"], cell_styles: %w[header header])
          s.row(["Item", 1234.56], cell_styles: [nil, "currency_cell"])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 3. Rich worksheet features (autofilter, data validation, conditional formatting, sparklines, etc.)
  def test_worksheet_rich_features_validation
    tmp = Tempfile.new(["xsd_rich_features", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("RichSheet", tab_color: "FFFF0000") do |s|
          s.column(0..1, width: 20)

          s.row(%w[Product Sales Status])
          s.row(["Widget A", 100, "Active"])
          s.row(["Widget B", 200, "Pending"])
          s.row(["Widget C", 300, "Active"])

          s.auto_filter("A1:C4")
          s.merge("A6:C6")
          s.freeze_pane(row: 1, column: 0)

          s.conditional_format("B2:B4",
                               type: :cellIs,
                               operator: :greaterThan,
                               formula: ["150"],
                               dxf: { font: { bold: true, color: "FF006100" }, fill: { fg_color: "FFC6EFCE" } })

          s.data_validation("C2:C4", type: :list, formula1: '"Active,Pending,Closed"')

          s.sparkline_group(type: :line,
                            sparklines: [{ location: "D2", sqref: "B2:B4" }])

          s.page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)
          s.page_setup(orientation: :landscape, paper_size: 9)
          s.header_footer(odd_header: "&L&G&CHeader Title", odd_footer: "&RPage &P of &N")
          s.sheet_protection(password: "secret", select_locked_cells: true)
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 4. Tables validation (table1.xml)
  def test_table_validation
    tmp = Tempfile.new(["xsd_table", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("TableSheet") do |s|
          s.row(%w[ID Name Department Salary])
          s.row([1, "Alice", "Engineering", 95_000])
          s.row([2, "Bob", "Design", 80_000])
          s.row([3, "Charlie", "Marketing", 75_000])

          s.table("A1:D4", columns: %w[ID Name Department Salary], name: "EmployeesTable", style: "TableStyleMedium9", total_row: false)
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  def test_table_calculated_columns_and_totals_row_formula_validation
    tmp = Tempfile.new(["xsd_table_formulas", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("FormulasTable") do |s|
          s.row(%w[Item Qty Price Total])
          s.row(["Widget", 10, 5.0, 50.0])
          s.row(["Gadget", 20, 15.0, 300.0])
          s.row(["Totals", nil, nil, 350.0])

          s.table("A1:D4") do |t|
            t.name = "TableWithFormulas"
            t.column(:item)
            t.column(:qty, totals_row_function: :sum)
            t.column(:price, totals_row_function: :average)
            t.calculated_column(:total, "[@Qty]*[@Price]", totals_row_formula: "SUBTOTAL(109,[Total])")
            t.totals_row = true
          end
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 5. Charts and Drawings validation (chart1.xml, drawing1.xml)
  def test_charts_and_drawings_validation
    tmp = Tempfile.new(["xsd_charts", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("ChartSheet") do |s|
          s.row(%w[Category Q1 Q2])
          s.row(["North", 100, 150])
          s.row(["South", 200, 250])
          s.row(["East", 120, 180])
          s.row(["West", 300, 320])

          s.chart(type: :bar,
                  title: "Regional Sales",
                  series: [
                    { cat_ref: "ChartSheet!$A$2:$A$5", val_ref: "ChartSheet!$B$2:$B$5" },
                    { cat_ref: "ChartSheet!$A$2:$A$5", val_ref: "ChartSheet!$C$2:$C$5" }
                  ],
                  legend: { position: "r" },
                  cat_axis_title: "Region",
                  val_axis_title: "Revenue")
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 6. Comments validation (comments1.xml)
  def test_comments_validation
    tmp = Tempfile.new(["xsd_comments", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("CommentSheet") do |s|
          s.row(["Reviewed Item", 100])
          s.comment("A1", "Approved by manager", author: "Auditor")
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 7. Document properties validation (app.xml, custom.xml)
  def test_document_properties_validation
    tmp = Tempfile.new(["xsd_docprops", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.properties(
          core: { title: "Financial Report", subject: "FY2026", creator: "Finance Dept" },
          app: { company: "Acme Corp", manager: "Jane Doe" },
          custom: { "Department" => "Accounting", "Reviewed" => true, "VersionCode" => 42 }
        )
        w.sheet("Data") do |s|
          s.row(["OK"])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 8. In-Memory Workbook build validation
  def test_in_memory_build_validation
    tmp = Tempfile.new(["xsd_in_memory", ".xlsx"])
    begin
      wb = Xlsxrb.build do |w|
        w.sheet("MemorySheet") do |s|
          s.row(["In-Memory", 999, Date.today])
          s.merge("A1:B1")
        end
      end
      Xlsxrb.write(tmp.path, wb)

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 9. All chart types validation (line, pie, area, scatter, doughnut, radar, bar3d)
  def test_all_chart_types_validation
    chart_types = %i[line pie area scatter doughnut radar bar3d]
    chart_types.each do |chart_type|
      tmp = Tempfile.new(["xsd_chart_#{chart_type}", ".xlsx"])
      begin
        Xlsxrb.write(tmp.path) do |w|
          w.sheet("ChartData") do |s|
            s.row(%w[Cat Val1 Val2])
            s.row(["A", 10, 20])
            s.row(["B", 30, 40])
            s.chart(type: chart_type,
                    title: "#{chart_type.to_s.capitalize} Chart",
                    series: [
                      { cat_ref: "ChartData!$A$2:$A$3", val_ref: "ChartData!$B$2:$B$3" }
                    ])
          end
        end

        validate_xlsx_file(tmp.path)
      ensure
        tmp.close!
      end
    end
  end

  # Combo chart with primary and secondary value axes validation
  def test_combo_chart_with_secondary_axis_xsd_validation
    tmp = Tempfile.new(["xsd_combo_chart", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Sales") do |s|
          s.row(%w[Month Revenue Units])
          s.row(["Jan", 10_000, 50])
          s.row(["Feb", 12_000, 65])
          s.row(["Mar", 15_000, 80])
          s.chart(type: :combo,
                  title: "Revenue vs Units",
                  category_axis: { title: "Month" },
                  value_axis: { title: "Revenue ($)" },
                  secondary_value_axis: { title: "Units Sold", major_gridlines: false },
                  series: [
                    { type: :bar, name: "Revenue", cat_ref: "Sales!$A$2:$A$4", val_ref: "Sales!$B$2:$B$4" },
                    { type: :line, name: "Units", cat_ref: "Sales!$A$2:$A$4", val_ref: "Sales!$C$2:$C$4", secondary_axis: true }
                  ])
        end
      end

      validate_xlsx_file(tmp.path)

      doc = Xlsxrb::Ooxml::Reader.new(tmp.path)
      charts = doc.charts(sheet: "Sales")
      assert_equal(1, charts.size)
      chart = charts.first
      assert_equal("comboChart", chart[:type])
      assert_equal(2, chart[:series].size)
      assert_equal("Revenue ($)", chart[:val_axis_title])
      assert_equal("Units Sold", chart[:sec_val_axis_title])
      assert_equal(true, chart[:series][1][:secondary_axis])
    ensure
      tmp.close!
    end
  end

  # 10. Shared strings rich text XML validation
  def test_rich_text_shared_strings_validation
    tmp = Tempfile.new(["xsd_rich_text", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("RichTextSheet") do |s|
          rt = Xlsxrb.rich_text(
            { text: "Bold Part", font: { bold: true, color: "FFFF0000" } },
            { text: " and " },
            { text: "Italic Part", font: { italic: true, color: "FF0000FF" } }
          )
          s.row([rt, "Normal String"])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  def test_rich_text_inline_strings_validation
    tmp = Tempfile.new(["xsd_inline_rich_text", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path, shared_strings: false) do |w|
        w.sheet("RichTextSheet") do |s|
          s.style(:header, bold: true)
          rt = Xlsxrb.rich_text(
            { text: "Bold Part", font: { bold: true, color: "FFFF0000" } },
            { text: " and " },
            { text: "Italic Part", font: { italic: true, color: "FF0000FF" } }
          )
          s.row([rt, "Normal String"])
          s.row([rt, "Styled String"], cell_styles: [:header, nil])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 11. Protected ranges and cell watches XML schema validation
  def test_protected_ranges_and_cell_watches_validation
    tmp = Tempfile.new(["xsd_prot_ranges_watches", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("ProtectedSheet") do |s|
          s.sheet_protection(sheet: true)
          s.protected_range("Range1", "A1:B10", password: "pwd")
          s.protected_range("Range2", "C1:D5")
          s.cell_watch("A1")
          s.cell_watch("C5")
          s.row(%w[Val1 Val2 Val3 Val4])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 12. Structured color (theme, indexed, auto) round-trip XML schema validation
  def test_structured_color_roundtrip_styles_validation
    tmp = Tempfile.new(["xsd_structured_colors", ".xlsx"])
    begin
      sheet = Xlsxrb::Elements::Worksheet.new(
        name: "Sheet1",
        rows: [
          Xlsxrb::Elements::Row.new(
            index: 0,
            cells: [
              Xlsxrb::Elements::Cell.fast_create(0, 0, "Test", 1, nil, nil, nil, false)
            ]
          )
        ],
        styles: {
          fonts: [
            {},
            { sz: 12, name: "Calibri", color: { theme: 1, tint: 0.4 } }
          ],
          fills: [
            { pattern: "none" },
            { pattern: "gray125" },
            { pattern: "solid", fg_color: { indexed: 64 }, bg_color: { theme: 2, tint: -0.25 } }
          ],
          borders: [
            {},
            { left: { style: "thin", color: { auto: true } }, right: { style: "double", color: { theme: 3 } } }
          ],
          cell_xfs: [
            { font_id: 0, fill_id: 0, border_id: 0, num_fmt_id: 0 },
            { font_id: 1, fill_id: 2, border_id: 1, num_fmt_id: 0 }
          ]
        }
      )
      wb = Xlsxrb::Elements::Workbook.new(sheets: [sheet])

      out = Xlsxrb.write(wb)
      tmp.binmode
      tmp.write(out)
      tmp.flush
      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 13. DrawingML line properties and character properties sequence order validation
  def test_drawingml_line_and_char_properties_sequence_order_validation
    tmp = Tempfile.new(["xsd_drawing_seq", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("Shapes") do |s|
          s.shape(
            preset: "roundRect",
            from_col: 1,
            from_row: 1,
            to_col: 5,
            to_row: 5,
            text: "Rounded",
            line_color: "FF0000",
            line_width: 2,
            line_join: "miter",
            line_miter_limit: 800_000,
            font: {
              name: "Arial",
              line_color: "0000FF",
              line_width: 1,
              outer_shadow: { blur_rad: 20_000, dist: 30_000, color: "000000" }
            }
          )
          s.row(["Cell"])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 14. Data validation with child sqref and page setup negative firstPageNumber validation
  def test_data_validation_and_page_setup_validation
    tmp = Tempfile.new(["xsd_dv_ps", ".xlsx"])
    begin
      sheet = Xlsxrb::Elements::Worksheet.new(
        name: "Sheet1",
        rows: [
          Xlsxrb::Elements::Row.new(index: 0, cells: [Xlsxrb::Elements::Cell.fast_create(0, 0, "A", 0, nil, nil, nil, false)])
        ],
        data_validations: [
          { sqref: "A1", type: "list", formula1: '"Option1,Option2"' },
          { sqref: nil, type: "list", formula1: "$B$1:$B$5" } # Should be omitted because sqref is required
        ],
        page_setup: {
          first_page_number: -1 # Should be omitted because firstPageNumber must be unsignedInt
        }
      )
      wb = Xlsxrb::Elements::Workbook.new(sheets: [sheet])

      out = Xlsxrb.write(wb)
      tmp.binmode
      tmp.write(out)
      tmp.flush
      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 15. Ignored errors XSD sequence conformance
  def test_ignored_errors_and_sheet_properties_xsd_validation
    tmp = Tempfile.new(["xsd_ignored_errors", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("IgnoredErrorsSheet") do |s|
          s.row(["0123", "=1/0"])
          s.ignored_error("A1:A10", number_stored_as_text: true)
          s.ignored_error("B1", eval_error: true, unlocked_formula: true)
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 16. Phonetic properties XSD sequence conformance
  def test_phonetic_properties_xsd_validation
    tmp = Tempfile.new(["xsd_phonetic", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("PhoneticSheet") do |s|
          s.row(%w[漢字 かんじ])
          s.phonetic_properties = { font_id: 0, type: :hiragana, alignment: :distributed }
          s.merge("A1:B1")
          s.conditional_format("A1:B1", type: :cell_is, operator: "notEqual", formula: ["\"\""], bold: true)
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 17. Code name, fit-to-page, and sheetCalcPr XSD sequence conformance
  def test_code_name_fit_to_page_and_sheet_calc_pr_xsd_validation
    tmp = Tempfile.new(["xsd_sheet_props", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.code_name = "ThisWorkbook"
        w.sheet("Sheet1", code_name: "Sheet1Code") do |s|
          s.row([1, 2, 3])
          s.fit_to_page = true
          s.page_setup(fit_to_width: 1, fit_to_height: 2)
          s.sheet_calc_properties = { full_calc_on_load: true }
          s.sheet_protection(password: "secret")
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 18. Background image (<picture>) XSD sequence conformance
  def test_background_image_xsd_validation
    tmp = Tempfile.new(["xsd_bg_image", ".xlsx"])
    dummy_png = [
      0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
      0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
      0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
      0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
      0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
      0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
      0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
      0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
      0x42, 0x60, 0x82
    ].pack("C*")
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("BgSheet") do |s|
          s.row([1, 2, 3])
          s.background_image = { data: dummy_png, ext: "png" }
          s.table("A1:C1", columns: %w[Col1 Col2 Col3], name: "Table1")
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 19. Extended conditional formatting rules (aboveAverage, top10, duplicateValues, containsErrors, timePeriod, etc.)
  def test_extended_conditional_formatting_xsd_validation
    tmp = Tempfile.new(["xsd_cf_extended", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("CfExtended") do |s|
          s.row([10, 20, 30, 40, 50])
          s.row([15, 25, 35, 45, 55])

          # 1. aboveAverage
          s.conditional_format("A1:A10") do |cf|
            cf.type = :above_average
            cf.above_average = true
            cf.equal_average = true
            cf.std_dev = 1
            cf.dxf = { font: { bold: true, color: "FF006100" } }
          end

          # 2. top10
          s.conditional_format("B1:B10") do |cf|
            cf.type = :top10
            cf.rank = 3
            cf.percent = false
            cf.bottom = true
            cf.dxf = { fill: { fg_color: "FFFFC7CE" } }
          end

          # 3. duplicateValues & uniqueValues
          s.conditional_format("C1:C10", type: :duplicate_values, dxf: { font: { italic: true } })
          s.conditional_format("C1:C10", type: :unique_values, dxf: { font: { color: "FF0000FF" } })

          # 4. containsBlanks & notContainsBlanks
          s.conditional_format("D1:D10", type: :contains_blanks, dxf: { fill: { fg_color: "FFFFFF00" } })
          s.conditional_format("D1:D10", type: :not_contains_blanks, dxf: { font: { bold: true } })

          # 5. containsErrors & notContainsErrors
          s.conditional_format("E1:E10", type: :contains_errors, dxf: { fill: { fg_color: "FFFF0000" } })
          s.conditional_format("E1:E10", type: :not_contains_errors, dxf: { font: { color: "FF00FF00" } })

          # 6. timePeriod
          s.conditional_format("F1:F10", time_period: :this_month, dxf: { font: { bold: true } })
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  # 20. Outline properties (<outlinePr>) and row/column grouping XSD sequence conformance
  def test_outlines_and_grouping_xsd_validation
    tmp = Tempfile.new(["xsd_outlines", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("OutlinesSheet") do |s|
          s.outline_properties(apply_styles: true, summary_below: true, summary_right: false, show_outline_symbols: true)
          s.group_columns(0..2, level: 1, collapsed: true)

          s.row(%w[A B C D])
          s.group_rows(level: 1, collapsed: true) do
            s.row([1, 2, 3, 4])
            s.row([5, 6, 7, 8])
          end
          s.row([6, 8, 10, 12])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  def test_font_decorations_and_vertical_alignment_xsd_validation
    tmp = Tempfile.new(["xsd_font_decorations", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.style("u_double", font: { underline: :double })
        w.style("u_single_acct", font: { underline: :single_accounting })
        w.style("u_double_acct", font: { underline: :double_accounting })
        w.style("super", font: { superscript: true })
        w.style("sub", font: { subscript: true })

        w.sheet("FontDecorations") do |s|
          s.row(["Double", "Single Acct", "Double Acct", "Super", "Sub"],
                cell_styles: %w[u_double u_single_acct u_double_acct super sub])

          rt = Xlsxrb.rich_text(
            { text: "E = mc", font: { bold: true } },
            { text: "2", font: { vert_align: :superscript, bold: true } },
            { text: " and H", font: { italic: true } },
            { text: "2", font: { vert_align: :subscript } },
            { text: "O with ", font: { underline: :double } },
            { text: "Accounting", font: { underline: :single_accounting } }
          )
          s.row([rt])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end

  def test_shared_formulas_xsd_validation
    tmp = Tempfile.new(["xsd_shared_formulas", ".xlsx"])
    begin
      Xlsxrb.write(tmp.path) do |w|
        w.sheet("SharedFormulasSheet") do |s|
          s.row([10, Xlsxrb.shared_formula("A1*2", shared_index: 0, ref: "B1:B3", cached_value: 20)])
          s.row([20, Xlsxrb.shared_formula(shared_index: 0, cached_value: 40)])
          s.row([30, { type: :shared, shared_index: 0, cached_value: 60 }])
        end
      end

      validate_xlsx_file(tmp.path)
    ensure
      tmp.close!
    end
  end
end
