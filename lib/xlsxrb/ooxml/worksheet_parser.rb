# frozen_string_literal: true

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # Streaming parser for xl/worksheets/sheetN.xml.
    # Uses a fast string-scanning approach for row/cell extraction,
    # falling back to REXML SAX only for column definitions and unmapped data.
    class WorksheetParser
      EMPTY_ARRAY = [].freeze
      EMPTY_HASH = {}.freeze

      # Parses all rows from a worksheet XML string.
      # Returns an Array of raw row hashes:
      #   { index:, cells: [{ ref:, type:, style_index:, value:, formula: }], attrs:, unmapped: }
      def self.parse(xml_string, shared_strings: [])
        return [] if xml_string.nil? || xml_string.empty?

        rows = []
        fast_scan_rows(xml_string, shared_strings) { |row| rows << row }
        rows
      end

      # Streaming parse: yields one raw row hash at a time.
      def self.each_row(xml_string, shared_strings: [], &block)
        return enum_for(:each_row, xml_string, shared_strings: shared_strings) unless block

        fast_scan_rows(xml_string, shared_strings, &block)
      end

      # Parses column definitions (<cols>) from a worksheet.
      def self.parse_columns(xml_string)
        return [] if xml_string.nil? || xml_string.empty?

        listener = ColumnsListener.new
        XmlParser.parse(xml_string, listener)
        listener.columns
      end

      # ---- Fast string-scanning parser (byte-level for O(1) offset) ----
      # All positions are byte offsets. We use byteindex/byteslice to avoid
      # O(n) character-offset conversion on UTF-8 strings.

      def self.fast_scan_rows(xml_src, shared_strings, &block)
        xml = xml_src.b # force ASCII-8BIT for O(1) byte indexing

        sd_start = xml.index("<sheetData")
        return unless sd_start

        sd_open_end = xml.index(">", sd_start)
        return unless sd_open_end

        return if xml.getbyte(sd_open_end - 1) == 47 # self-closing <sheetData/>

        sd_end = xml.index("</sheetData>", sd_open_end)
        return unless sd_end

        pos = sd_open_end + 1

        while pos < sd_end
          row_start = xml.index("<row", pos)
          break unless row_start && row_start < sd_end

          nb = xml.getbyte(row_start + 4)
          unless [32, 62, 9, 10, 13, 47].include?(nb)
            pos = row_start + 4
            next
          end

          tag_end = xml.index(">", row_start + 4)
          break unless tag_end

          if xml.getbyte(tag_end - 1) == 47
            pos = tag_end + 1
            next
          end

          # Row index and attrs from tag substring (bounded search)
          row_tag = xml.byteslice(row_start, tag_end - row_start)
          row_index = 0
          r_val = tag_attr(row_tag, ' r="')
          row_index = r_val.to_i - 1 if r_val

          attrs = extract_row_attrs(row_tag)

          row_end = xml.index("</row>", tag_end + 1)
          break unless row_end

          cells = fast_parse_cells(xml, tag_end + 1, row_end, shared_strings)

          block.call({ index: row_index, cells: cells, attrs: attrs, unmapped: EMPTY_ARRAY })

          pos = row_end + 6
        end
      end

      private_class_method :fast_scan_rows

      def self.extract_row_attrs(row_tag)
        attrs = EMPTY_HASH

        ht_val = tag_attr(row_tag, ' ht="')
        if ht_val
          attrs = {}
          attrs[:height] = ht_val.to_f
        end

        if row_tag.include?('hidden="1"')
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:hidden] = true
        end

        if row_tag.include?('customHeight="1"')
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:custom_height] = true
        end

        ol_val = tag_attr(row_tag, ' outlineLevel="')
        if ol_val
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:outline_level] = ol_val.to_i
        end

        attrs
      end

      private_class_method :extract_row_attrs

      # Extract an attribute value from a small tag substring (bounded search).
      def self.tag_attr(tag, prefix)
        a_pos = tag.index(prefix)
        return nil unless a_pos

        val_start = a_pos + prefix.bytesize
        val_end = tag.index('"', val_start)
        return nil unless val_end

        tag.byteslice(val_start, val_end - val_start).force_encoding("UTF-8")
      end

      private_class_method :tag_attr

      def self.fast_parse_cells(xml, from, to, shared_strings)
        cells = []
        pos = from

        while pos < to
          c_start = xml.index("<c", pos)
          break unless c_start && c_start < to

          nb = xml.getbyte(c_start + 2)
          unless [32, 62, 9, 10, 13, 47].include?(nb)
            pos = c_start + 2
            next
          end

          c_tag_end = xml.index(">", c_start + 2)
          break unless c_tag_end

          # Extract tag substring for bounded attribute search
          c_tag = xml.byteslice(c_start, c_tag_end - c_start)
          ref = tag_attr(c_tag, ' r="')
          type = tag_attr(c_tag, ' t="')
          style_str = tag_attr(c_tag, ' s="')
          style_index = style_str&.to_i

          # Self-closing <c ... />
          if xml.getbyte(c_tag_end - 1) == 47
            cells << { ref: ref, type: type, style_index: style_index, value: nil }
            pos = c_tag_end + 1
            next
          end

          c_end = xml.index("</c>", c_tag_end + 1)
          break unless c_end

          # Parse cell content sequentially (bounded to c_end - avoid unbounded scans)
          value = nil
          formula = nil
          inline_str = nil
          cpos = c_tag_end + 1
          while cpos < c_end
            tag_pos = xml.index("<", cpos)
            break unless tag_pos && tag_pos < c_end

            tag_char = xml.getbyte(tag_pos + 1)
            case tag_char
            when 118 # 'v'
              if xml.getbyte(tag_pos + 2) == 62 # <v>
                v_val_start = tag_pos + 3
                v_end = xml.index("</v>", v_val_start)
                if v_end
                  raw_value = xml.byteslice(v_val_start, v_end - v_val_start)
                  value = resolve_fast_value(raw_value, type, shared_strings)
                  cpos = v_end + 4
                else
                  cpos = tag_pos + 3
                end
              elsif xml.getbyte(tag_pos + 2) == 47 && xml.getbyte(tag_pos + 3) == 62 # <v/>
                cpos = tag_pos + 4
              else
                cpos = tag_pos + 2
              end
            when 102 # 'f'
              f_tag_end = xml.index(">", tag_pos + 2)
              if f_tag_end && f_tag_end < c_end
                if xml.getbyte(f_tag_end - 1) == 47 # self-closing <f ... />
                  cpos = f_tag_end + 1
                else
                  f_end = xml.index("</f>", f_tag_end + 1)
                  if f_end && f_end <= c_end
                    formula = xml.byteslice(f_tag_end + 1, f_end - f_tag_end - 1).force_encoding("UTF-8")
                    formula = decode_xml_entities(formula) if formula.include?("&")
                    cpos = f_end + 4
                  else
                    cpos = f_tag_end + 1
                  end
                end
              else
                cpos = tag_pos + 2
              end
            when 105 # 'i' - <is>
              if xml.byteslice(tag_pos, 4) == "<is>"
                is_end = xml.index("</is>", tag_pos + 4)
                if is_end && is_end <= c_end
                  inline_str = extract_inline_text(xml, tag_pos + 4, is_end)
                  cpos = is_end + 5
                else
                  cpos = tag_pos + 4
                end
              else
                cpos = tag_pos + 2
              end
            else
              # Skip unknown tag
              close = xml.index(">", tag_pos + 1)
              cpos = close ? close + 1 : c_end
            end
          end

          cell = { ref: ref, type: type, style_index: style_index, value: value }
          cell[:formula] = formula if formula
          cell[:inline_string] = inline_str if inline_str
          cells << cell

          pos = c_end + 4
        end

        cells
      end

      private_class_method :fast_parse_cells

      def self.extract_inline_text(xml, from, to)
        result = +""
        pos = from
        while pos < to
          t_start = xml.index("<t", pos)
          break unless t_start && t_start < to

          t_tag_end = xml.index(">", t_start)
          break unless t_tag_end
          next (pos = t_start + 2) if xml.getbyte(t_tag_end - 1) == 47

          t_end = xml.index("</t>", t_tag_end + 1)
          break unless t_end && t_end <= to

          result << xml.byteslice(t_tag_end + 1, t_end - t_tag_end - 1)
          pos = t_end + 4
        end
        result.force_encoding("UTF-8")
        result = decode_xml_entities(result) if result.include?("&")
        result
      end

      private_class_method :extract_inline_text

      def self.resolve_fast_value(raw, type, shared_strings)
        case type
        when "s"
          shared_strings[raw.to_i] || ""
        when "b"
          raw == "1"
        when "e", "str", "inlineStr"
          val = raw.force_encoding("UTF-8")
          val.include?("&") ? decode_xml_entities(val) : val
        else
          return nil if raw.empty?

          if raw.include?(".")
            raw.to_f
          else
            int_val = raw.to_i
            int_val.to_s == raw ? int_val : raw.to_f
          end
        end
      end

      private_class_method :resolve_fast_value

      XML_ENTITIES = { "&amp;" => "&", "&lt;" => "<", "&gt;" => ">", "&quot;" => '"', "&apos;" => "'" }.freeze

      def self.decode_xml_entities(str)
        str.gsub(/&(?:amp|lt|gt|quot|apos);/, XML_ENTITIES)
      end

      private_class_method :decode_xml_entities

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
