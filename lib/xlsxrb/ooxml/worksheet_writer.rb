# frozen_string_literal: true

require "stringio"
require_relative "xml_builder"

module Xlsxrb
  module Ooxml
    # Generates worksheet XML for a list of rows.
    # Supports streaming: rows can be written one at a time.
    class WorksheetWriter
      SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

      def initialize(io)
        @io = io
        @builder = XmlBuilder.new(@io)
        @started = false
        @finished = false
      end

      # Write the worksheet header. Call once before writing rows.
      def start(columns: [])
        return if @started

        @started = true
        @builder.declaration
        @builder.open_tag("worksheet", { xmlns: SSML_NS })

        write_columns(columns) unless columns.empty?

        @builder.open_tag("sheetData")
      end

      # Write a single row. Automatically calls start if needed.
      def write_row(row_index, cells, attrs: {}, unmapped: [])
        start unless @started

        row_attrs = { r: (row_index + 1).to_s } # convert 0-based to 1-based
        row_attrs[:ht] = attrs[:height].to_s if attrs[:height]
        row_attrs[:hidden] = "1" if attrs[:hidden]
        row_attrs[:customHeight] = "1" if attrs[:custom_height]

        @builder.open_tag("row", row_attrs)

        cells.each do |cell|
          write_cell(cell)
        end

        unmapped.each { |node| @builder.write_unmapped(node) }

        @builder.close_tag("row")
      end

      # Write the worksheet footer. Call once after all rows.
      def finish
        return if @finished

        start unless @started
        @finished = true
        @builder.close_tag("sheetData")
        @builder.close_tag("worksheet")
      end

      private

      def write_columns(columns)
        @builder.open_tag("cols")
        columns.each do |col|
          attrs = {
            min: ((col[:index] || col[:min] || 0) + 1).to_s,
            max: ((col[:index] || col[:max] || col[:min] || 0) + 1).to_s
          }
          attrs[:width] = col[:width].to_s if col[:width]
          attrs[:hidden] = "1" if col[:hidden]
          attrs[:customWidth] = "1" if col[:custom_width] || col[:width]
          @builder.empty_tag("col", attrs)
        end
        @builder.close_tag("cols")
      end

      def write_cell(cell)
        ref = cell[:ref] || cell_ref(cell[:row_index], cell[:column_index])
        attrs = { r: ref }

        value = cell[:value]
        type = cell[:type] || cell_type(value)
        attrs[:t] = type if type
        attrs[:s] = cell[:style_index].to_s if cell[:style_index]

        formula = cell[:formula]

        if value.nil? && formula.nil?
          @builder.empty_tag("c", attrs)
          return
        end

        @builder.open_tag("c", attrs)
        @builder.tag("f") { |b| b.text(formula) } if formula
        @builder.tag("v") { |b| b.text(xml_cell_value(value, type)) } unless value.nil?
        @builder.close_tag("c")
      end

      def cell_type(value)
        case value
        when String then "s" # will be shared string index
        when true, false then "b"
        end
      end

      def xml_cell_value(value, _type)
        case value
        when true then "1"
        when false then "0"
        else value.to_s
        end
      end

      def cell_ref(row_index, col_index)
        col_letter = column_letter(col_index)
        "#{col_letter}#{row_index + 1}"
      end

      def column_letter(index)
        result = +""
        i = index
        loop do
          result.prepend(("A".ord + (i % 26)).chr)
          i = (i / 26) - 1
          break if i.negative?
        end
        result
      end
    end
  end
end
