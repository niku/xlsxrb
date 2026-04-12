# frozen_string_literal: true

require "stringio"
require_relative "xml_builder"
require_relative "zip_writer"
require_relative "worksheet_writer"

module Xlsxrb
  module Ooxml
    # Orchestrates writing a complete XLSX workbook.
    # Generates all required OpenXML parts and assembles them into a ZIP.
    class WorkbookWriter
      SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      REL_NS  = "http://schemas.openxmlformats.org/package/2006/relationships"
      CT_NS   = "http://schemas.openxmlformats.org/package/2006/content-types"
      DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

      def self.write(target, sheets:, shared_strings: [], styles: nil)
        writer = new(sheets: sheets, shared_strings: shared_strings, styles: styles)
        writer.write_to(target)
      end

      def initialize(sheets:, shared_strings: [], styles: nil)
        @sheets = sheets
        @shared_strings = shared_strings
        @styles = styles
        @drawing_count = 0
        @chart_count = 0
      end

      def write_to(target)
        ZipWriter.open(target) do |zip|
          zip.add_entry("[Content_Types].xml", build_content_types)
          zip.add_entry("_rels/.rels", build_root_rels)
          zip.add_entry("xl/workbook.xml", build_workbook_xml)
          zip.add_entry("xl/_rels/workbook.xml.rels", build_workbook_rels)
          zip.add_entry("xl/styles.xml", build_styles_xml)
          zip.add_entry("xl/sharedStrings.xml", build_shared_strings_xml) unless @shared_strings.empty?

          @sheets.each_with_index do |sheet, idx|
            drawing_rid = nil
            if sheet[:charts] && !sheet[:charts].empty?
              @drawing_count += 1
              drawing_idx = @drawing_count
              drawing_rels_data = []
              drawing_parts = []

              chart_writer = Xlsxrb::Ooxml::Writer.new
              chart_writer.add_sheet(sheet[:name])

              # Populate sheet data in chart_writer so chart cache can be resolved
              (sheet[:rows] || []).each do |row|
                (row[:cells] || []).each do |cell|
                  chart_writer.set_cell(cell[:ref], cell[:value], sheet: sheet[:name])
                end
              end

              sheet[:charts].each do |chart_options|
                chart_writer.add_chart(**chart_options)
              end

              processed_charts = chart_writer.charts

              processed_charts.each do |chart|
                @chart_count += 1
                chart_path = "xl/charts/chart#{@chart_count}.xml"
                zip.add_entry(chart_path, chart_writer.send(:generate_chart_xml, chart))
                drawing_rels_data << { type: :chart, target: "../charts/chart#{@chart_count}.xml" }
                drawing_parts << { kind: :chart, chart: chart, rid_index: drawing_rels_data.size }
              end

              drawing_xml = chart_writer.send(:generate_drawing_xml, drawing_parts)
              zip.add_entry("xl/drawings/drawing#{drawing_idx}.xml", drawing_xml)

              unless drawing_rels_data.empty?
                drawing_rels_xml = chart_writer.send(:generate_drawing_rels, drawing_rels_data)
                zip.add_entry("xl/drawings/_rels/drawing#{drawing_idx}.xml.rels", drawing_rels_xml)
              end

              # Assuming drawing is the only relationship for now
              drawing_rid = "rId1"
              sheet_rels_xml = build_sheet_rels(drawing_idx)
              zip.add_entry("xl/worksheets/_rels/sheet#{idx + 1}.xml.rels", sheet_rels_xml)
            end

            zip.add_entry("xl/worksheets/sheet#{idx + 1}.xml", build_worksheet_xml(sheet, drawing_rid: drawing_rid))
          end
        end
      end

      private

      def build_sheet_rels(drawing_idx)
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Relationships", { xmlns: REL_NS })
        b.empty_tag("Relationship", {
                      Id: "rId1",
                      Type: "#{DOC_REL}/drawing",
                      Target: "../drawings/drawing#{drawing_idx}.xml"
                    })
        b.close_tag("Relationships")
        io.string
      end

      def build_content_types
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Types", { xmlns: CT_NS })
        b.empty_tag("Default", { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" })
        b.empty_tag("Default", { Extension: "xml", ContentType: "application/xml" })
        b.empty_tag("Override", { PartName: "/xl/workbook.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" })
        b.empty_tag("Override", { PartName: "/xl/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" })
        b.empty_tag("Override", { PartName: "/xl/sharedStrings.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" }) unless @shared_strings.empty?

        drawing_count = 0
        chart_count = 0
        @sheets.each_with_index do |sheet, idx|
          b.empty_tag("Override", { PartName: "/xl/worksheets/sheet#{idx + 1}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" })
          next unless sheet[:charts] && !sheet[:charts].empty?

          drawing_count += 1
          b.empty_tag("Override", { PartName: "/xl/drawings/drawing#{drawing_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml" })
          sheet[:charts].each do |_chart|
            chart_count += 1
            b.empty_tag("Override", { PartName: "/xl/charts/chart#{chart_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" })
          end
        end
        b.close_tag("Types")
        io.string
      end

      def build_root_rels
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Relationships", { xmlns: REL_NS })
        b.empty_tag("Relationship", {
                      Id: "rId1",
                      Type: "#{DOC_REL}/officeDocument",
                      Target: "xl/workbook.xml"
                    })
        b.close_tag("Relationships")
        io.string
      end

      def build_workbook_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("workbook", {
                     xmlns: SSML_NS,
                     "xmlns:r": DOC_REL
                   })
        b.open_tag("sheets")
        @sheets.each_with_index do |sheet, idx|
          b.empty_tag("sheet", {
                        name: sheet[:name],
                        sheetId: (idx + 1).to_s,
                        "r:id": "rId#{idx + 1}"
                      })
        end
        b.close_tag("sheets")
        b.close_tag("workbook")
        io.string
      end

      def build_workbook_rels
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Relationships", { xmlns: REL_NS })
        @sheets.each_with_index do |_sheet, idx|
          b.empty_tag("Relationship", {
                        Id: "rId#{idx + 1}",
                        Type: "#{DOC_REL}/worksheet",
                        Target: "worksheets/sheet#{idx + 1}.xml"
                      })
        end
        rid = @sheets.size + 1
        b.empty_tag("Relationship", {
                      Id: "rId#{rid}",
                      Type: "#{DOC_REL}/styles",
                      Target: "styles.xml"
                    })
        unless @shared_strings.empty?
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "#{DOC_REL}/sharedStrings",
                        Target: "sharedStrings.xml"
                      })
        end
        b.close_tag("Relationships")
        io.string
      end

      def build_styles_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("styleSheet", { xmlns: SSML_NS })

        # Fonts
        fonts = @styles&.dig(:fonts) || []
        fonts = [{}] if fonts.empty? # At least one default font
        b.tag("fonts", { count: fonts.size.to_s }) do |_|
          fonts.each do |font_props|
            b.tag("font") do |_|
              b.empty_tag("sz", { val: (font_props[:sz] || 11).to_s }) if font_props[:sz] || fonts == [{}]
              b.empty_tag("bold") if font_props[:bold]
              b.empty_tag("italic") if font_props[:italic]
              b.empty_tag("strike") if font_props[:strike]
              if font_props[:underline]
                b.empty_tag("u", font_props[:underline] == "single" ? {} : { val: font_props[:underline] })
              end
              b.empty_tag("color", { rgb: font_props[:color] }) if font_props[:color]
              b.empty_tag("name", { val: font_props[:name] || "Calibri" })
            end
          end
        end

        # Fills
        fills = @styles&.dig(:fills) || []
        fills = [{ pattern: "none" }, { pattern: "gray125" }] if fills.empty?
        b.tag("fills", { count: fills.size.to_s }) do |_|
          fills.each do |fill_props|
            b.tag("fill") do |_|
              if fill_props[:gradient]
                b.open_tag("gradientFill", { type: fill_props[:gradient][:type], degree: fill_props[:gradient][:degree]&.to_s }.compact)
                fill_props[:gradient][:stops].each do |stop|
                  b.empty_tag("stop", { position: stop[:position].to_s })
                  b.empty_tag("color", { rgb: stop[:color] })
                end
                b.close_tag("gradientFill")
              else
                b.tag("patternFill", { patternType: fill_props[:pattern] || "none" }) do |_|
                  b.empty_tag("fgColor", { rgb: fill_props[:fg_color] }) if fill_props[:fg_color]
                  b.empty_tag("bgColor", { rgb: fill_props[:bg_color] }) if fill_props[:bg_color]
                end
              end
            end
          end
        end

        # Borders
        borders = @styles&.dig(:borders) || []
        borders = [{}] if borders.empty?
        b.tag("borders", { count: borders.size.to_s }) do |_|
          borders.each do |border_props|
            border_attrs = {}
            border_attrs[:diagonalUp] = "1" if border_props[:diagonal_up]
            border_attrs[:diagonalDown] = "1" if border_props[:diagonal_down]
            b.tag("border", border_attrs) do |_|
              %i[left right top bottom diagonal].each do |side|
                side_data = border_props[side]
                if side_data
                  b.tag(side.to_s, { style: side_data[:style] }) do |_|
                    b.empty_tag("color", { rgb: side_data[:color] }) if side_data[:color]
                  end
                else
                  b.empty_tag(side.to_s)
                end
              end
            end
          end
        end

        # cellStyleXfs
        b.tag("cellStyleXfs", { count: "1" }) do |_|
          b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" })
        end

        # Number formats
        num_fmts = @styles&.dig(:num_fmts) || []
        unless num_fmts.empty?
          b.tag("numFmts", { count: num_fmts.size.to_s }) do |_|
            num_fmts.each do |nf|
              if nf.is_a?(Hash)
                b.empty_tag("numFmt", { numFmtId: nf[:num_fmt_id].to_s, formatCode: nf[:format_code] })
              else
                # Legacy format handling
              end
            end
          end
        end

        # Cell XFs (formatting)
        xf_entries = @styles&.dig(:xf_entries) || []
        xf_count = [xf_entries.size, 1].max
        b.tag("cellXfs", { count: xf_count.to_s }) do |_|
          xf_entries.each do |xf|
            attrs = {
              numFmtId: (xf[:num_fmt_id] || 0).to_s,
              fontId: (xf[:font_id] || 0).to_s,
              fillId: (xf[:fill_id] || 0).to_s,
              borderId: (xf[:border_id] || 0).to_s
            }
            attrs[:xfId] = xf[:xf_id].to_s if xf[:xf_id]
            b.empty_tag("xf", attrs)
          end
          # Add default if empty
          b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" }) if xf_entries.empty?
        end

        b.close_tag("styleSheet")
        io.string
      end

      def build_shared_strings_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("sst", {
                     xmlns: SSML_NS,
                     count: @shared_strings.size.to_s,
                     uniqueCount: @shared_strings.size.to_s
                   })
        @shared_strings.each do |str|
          b.tag("si") { |_| b.tag("t") { |_| b.text(str) } }
        end
        b.close_tag("sst")
        io.string
      end

      def build_worksheet_xml(sheet, drawing_rid: nil)
        io = StringIO.new
        ws = WorksheetWriter.new(io)
        ws.start(columns: sheet[:columns] || [])
        (sheet[:rows] || []).each do |row|
          ws.write_row(row[:index], row[:cells], attrs: row[:attrs] || {}, unmapped: row[:unmapped] || [])
        end
        ws.finish(drawing_rid: drawing_rid)
        io.string
      end
    end
  end
end
