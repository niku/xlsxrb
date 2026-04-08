# frozen_string_literal: true

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # SAX-based streaming parser for xl/worksheets/sheetN.xml.
    # Yields raw row data one row at a time for memory-efficient processing.
    class WorksheetParser
      # Parses all rows from a worksheet XML string.
      # Returns an Array of raw row hashes:
      #   { index:, cells: [{ ref:, type:, style_index:, value:, formula: }], attrs:, unmapped: }
      def self.parse(xml_string, shared_strings: [])
        return [] if xml_string.nil? || xml_string.empty?

        listener = Listener.new(shared_strings)
        XmlParser.parse(xml_string, listener)
        listener.rows
      end

      # Streaming parse: yields one raw row hash at a time.
      def self.each_row(xml_string, shared_strings: [], &block)
        return enum_for(:each_row, xml_string, shared_strings: shared_strings) unless block

        listener = StreamingListener.new(shared_strings, &block)
        XmlParser.parse(xml_string, listener)
      end

      # Parses column definitions (<cols>) from a worksheet.
      def self.parse_columns(xml_string)
        return [] if xml_string.nil? || xml_string.empty?

        listener = ColumnsListener.new
        XmlParser.parse(xml_string, listener)
        listener.columns
      end

      # Shared row-parsing logic.
      module RowParsing
        RECOGNIZED_TAGS = %w[worksheet sheetData row c v f is t cols col sheetViews sheetView
                             dimension sheetFormatPr mergeCell mergeCells].freeze

        def init_row_state(shared_strings)
          @shared_strings = shared_strings
          @in_row = false
          @in_cell = false
          @in_value = false
          @in_formula = false
          @in_inline_string = false
          @in_t = false
          @current_row_index = nil
          @current_row_attrs = {}
          @current_row_cells = []
          @current_cell = nil
          @current_text = +""
          @current_formula = +""
          @current_inline = +""
          @unmapped_stack = []
          @capturing_unmapped = false
          @current_row_unmapped = []
        end

        def handle_start(localname, attrs)
          if @capturing_unmapped
            child = { tag: localname, attrs: attrs.dup, children: [], text: nil }
            @unmapped_stack.last[:children] << child
            @unmapped_stack.push(child)
            return
          end

          case localname
          when "row"
            @in_row = true
            @current_row_index = (attrs["r"]&.to_i || 1) - 1 # convert to 0-based
            @current_row_attrs = {}
            @current_row_attrs[:height] = attrs["ht"]&.to_f if attrs["ht"]
            @current_row_attrs[:hidden] = true if attrs["hidden"] == "1"
            @current_row_attrs[:custom_height] = true if attrs["customHeight"] == "1"
            @current_row_attrs[:outline_level] = attrs["outlineLevel"]&.to_i if attrs["outlineLevel"]
            @current_row_cells = []
            @current_row_unmapped = []
          when "c"
            @in_cell = true
            @current_cell = {
              ref: attrs["r"],
              type: attrs["t"],
              style_index: attrs["s"]&.to_i
            }
            @current_text = +""
            @current_formula = +""
            @current_inline = +""
          when "v"
            @in_value = true
            @current_text = +""
          when "f"
            @in_formula = true
            @current_formula = +""
          when "is"
            @in_inline_string = true
            @current_inline = +""
          when "t"
            @in_t = true
          else
            return unless @in_row

            # Unknown tag inside row → capture as unmapped
            @capturing_unmapped = true
            node = { tag: localname, attrs: attrs.dup, children: [], text: nil }
            @unmapped_stack.push(node)
          end
        end

        def handle_end(localname)
          if @capturing_unmapped
            @unmapped_stack.pop
            @capturing_unmapped = false if @unmapped_stack.empty?
            return
          end

          case localname
          when "v"
            @in_value = false
          when "f"
            @in_formula = false
          when "t"
            @in_t = false
          when "is"
            @in_inline_string = false
          when "c"
            finalize_cell
            @in_cell = false
          when "row"
            finalize_row
            @in_row = false
          end
        end

        def handle_characters(text)
          if @capturing_unmapped && !@unmapped_stack.empty?
            current = @unmapped_stack.last
            current[:text] = (current[:text] || "") + text
            return
          end

          if @in_value
            @current_text << text
          elsif @in_formula
            @current_formula << text
          elsif @in_inline_string && @in_t
            @current_inline << text
          end
        end

        private

        def finalize_cell
          return unless @current_cell

          cell = @current_cell
          raw_value = @current_text
          cell[:value] = resolve_cell_value(raw_value, cell[:type])
          cell[:formula] = @current_formula unless @current_formula.empty?
          cell[:inline_string] = @current_inline unless @current_inline.empty?
          @current_row_cells << cell
          @current_cell = nil
        end

        def resolve_cell_value(raw, type)
          case type
          when "s" # shared string
            idx = raw.to_i
            @shared_strings[idx] || ""
          when "b" # boolean
            raw == "1"
          when "e", "str", "inlineStr" # error / formula string / inline
            raw
          else
            # Numeric or general
            return nil if raw.empty?

            if raw.include?(".")
              raw.to_f
            else
              int_val = raw.to_i
              int_val.to_s == raw ? int_val : raw.to_f
            end
          end
        end
      end

      # Collects all rows in memory.
      class Listener
        include REXML::SAX2Listener
        include RowParsing

        attr_reader :rows

        def initialize(shared_strings)
          init_row_state(shared_strings)
          @rows = []
        end

        def start_element(_uri, localname, _qname, attrs)
          handle_start(localname, attrs)
        end

        def end_element(_uri, localname, _qname)
          handle_end(localname)
        end

        def characters(text)
          handle_characters(text)
        end

        private

        def finalize_row
          @rows << {
            index: @current_row_index,
            cells: @current_row_cells.dup,
            attrs: @current_row_attrs.dup,
            unmapped: @current_row_unmapped.dup
          }
        end
      end

      # Yields each row to a block (streaming).
      class StreamingListener
        include REXML::SAX2Listener
        include RowParsing

        def initialize(shared_strings, &block)
          init_row_state(shared_strings)
          @block = block
        end

        def start_element(_uri, localname, _qname, attrs)
          handle_start(localname, attrs)
        end

        def end_element(_uri, localname, _qname)
          handle_end(localname)
        end

        def characters(text)
          handle_characters(text)
        end

        private

        def finalize_row
          row_data = {
            index: @current_row_index,
            cells: @current_row_cells.dup,
            attrs: @current_row_attrs.dup,
            unmapped: @current_row_unmapped.dup
          }
          @block.call(row_data)
        end
      end

      # Parses <cols> section for column definitions.
      class ColumnsListener
        include REXML::SAX2Listener

        attr_reader :columns

        def initialize
          @columns = []
          @in_cols = false
        end

        def start_element(_uri, localname, _qname, attrs)
          case localname
          when "cols"
            @in_cols = true
          when "col"
            return unless @in_cols

            col = {
              min: attrs["min"]&.to_i,
              max: attrs["max"]&.to_i,
              width: attrs["width"]&.to_f,
              hidden: attrs["hidden"] == "1",
              custom_width: attrs["customWidth"] == "1",
              outline_level: attrs["outlineLevel"]&.to_i
            }
            @columns << col
          end
        end

        def end_element(_uri, localname, _qname)
          @in_cols = false if localname == "cols"
        end

        def characters(_text); end
      end
    end
  end
end
