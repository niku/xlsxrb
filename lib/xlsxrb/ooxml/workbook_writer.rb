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
            zip.add_entry("xl/worksheets/sheet#{idx + 1}.xml", build_worksheet_xml(sheet))
          end
        end
      end

      private

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
        @sheets.each_with_index do |_sheet, idx|
          b.empty_tag("Override", { PartName: "/xl/worksheets/sheet#{idx + 1}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" })
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

        # Minimal default styles
        b.tag("fonts", { count: "1" }) do |_|
          b.tag("font") do |_|
            b.empty_tag("sz", { val: "11" })
            b.empty_tag("name", { val: "Calibri" })
          end
        end
        b.tag("fills", { count: "2" }) do |_|
          b.tag("fill") { |_| b.empty_tag("patternFill", { patternType: "none" }) }
          b.tag("fill") { |_| b.empty_tag("patternFill", { patternType: "gray125" }) }
        end
        b.tag("borders", { count: "1" }) do |_|
          b.tag("border") do |_|
            b.empty_tag("left")
            b.empty_tag("right")
            b.empty_tag("top")
            b.empty_tag("bottom")
            b.empty_tag("diagonal")
          end
        end
        b.tag("cellStyleXfs", { count: "1" }) do |_|
          b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" })
        end

        num_fmts = @styles&.dig(:num_fmts) || {}
        xf_count = [@styles&.dig(:cell_xfs)&.size || 1, 1].max

        unless num_fmts.empty?
          b.tag("numFmts", { count: num_fmts.size.to_s }) do |_|
            num_fmts.each do |id, code|
              b.empty_tag("numFmt", { numFmtId: id.to_s, formatCode: code })
            end
          end
        end

        b.tag("cellXfs", { count: xf_count.to_s }) do |_|
          if @styles&.dig(:cell_xfs)
            @styles[:cell_xfs].each do |xf|
              attrs = { numFmtId: (xf[:num_fmt_id] || 0).to_s, fontId: (xf[:font_id] || 0).to_s,
                        fillId: (xf[:fill_id] || 0).to_s, borderId: (xf[:border_id] || 0).to_s }
              b.empty_tag("xf", attrs)
            end
          else
            b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" })
          end
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

      def build_worksheet_xml(sheet)
        io = StringIO.new
        ws = WorksheetWriter.new(io)
        ws.start(columns: sheet[:columns] || [])
        (sheet[:rows] || []).each do |row|
          ws.write_row(row[:index], row[:cells], attrs: row[:attrs] || {}, unmapped: row[:unmapped] || [])
        end
        ws.finish
        io.string
      end
    end
  end
end
