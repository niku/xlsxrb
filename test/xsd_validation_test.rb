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
          s.row(["Header 1", "Header 2"], styles: %w[header header])
          s.row(["Item", 1234.56], styles: [nil, "currency_cell"])
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
          s.freeze_pane(row: 1, col: 0)

          s.conditional_format("B2:B4",
                               type: :cellIs,
                               operator: :greaterThan,
                               formula: ["150"],
                               dxf: { font: { bold: true, color: "FF006100" }, fill: { fg_color: "FFC6EFCE" } })

          s.validate_data("C2:C4", type: :list, formula1: '"Active,Pending,Closed"')

          s.sparkline_group(type: :line,
                            sparklines: [{ location: "D2", sqref: "B2:B4" }])

          s.page_margins(left: 0.7, right: 0.7, top: 0.75, bottom: 0.75)
          s.page_setup(orientation: :landscape, paper_size: 9)
          s.header_footer(odd_header: "&L&G&CHeader Title", odd_footer: "&RPage &P of &N")
          s.protect_sheet(password: "secret", select_locked_cells: true)
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
end
