# frozen_string_literal: true

# rbs_inline: enabled

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
        each_row(xml_string, shared_strings: shared_strings) { |row| rows << row }
        rows
      end

      # Streaming parse: yields one raw row hash at a time.
      def self.each_row(xml_string, shared_strings: [], part_name: "xl/worksheets/sheet1.xml", &block)
        return enum_for(:each_row, xml_string, shared_strings: shared_strings, part_name: part_name) unless block
        return if xml_string.nil? || xml_string.empty?

        fast_scan_rows_direct(xml_string, shared_strings, part_name, &block)
      end

      # Parses column definitions (<cols>) from a worksheet.
      def self.parse_columns(xml_string, part_name: "xl/worksheets/sheet1.xml")
        return [] if xml_string.nil? || xml_string.empty?

        columns = []
        each_event(xml_string, part_name: part_name) do |event|
          next unless event.type == :column

          min, max, width, hidden, custom_width, outline_level = event.args
          columns << {
            min: min,
            max: max,
            width: width,
            hidden: hidden,
            custom_width: custom_width,
            outline_level: outline_level
          }
        end
        columns
      end

      # Yields Event objects for columns, rows, cells, and hyperlinks.
      def self.each_event(xml_string, shared_strings: [], part_name: "xl/worksheets/sheet1.xml", &block)
        return enum_for(:each_event, xml_string, shared_strings: shared_strings, part_name: part_name) unless block
        return if xml_string.nil? || xml_string.empty?

        # 1. Parse and yield column events
        if xml_string.include?("<cols")
          listener = ColumnsListener.new
          XmlParser.parse(xml_string, listener)
          listener.columns.each do |col|
            block.call(Event.new(
                         type: :column,
                         args: [col[:min], col[:max], col[:width], col[:hidden], col[:custom_width], col[:outline_level]],
                         source: { part: part_name }
                       ))
          end
        end

        # 2. Parse and yield row/cell events
        fast_scan_events(xml_string, shared_strings, part_name, &block)

        # 3. Parse and yield hyperlink events
        return unless xml_string.include?("hyperlink")

        xml = xml_string.b
        hpos_term = xml.index("hyperlinks")
        return unless hpos_term

        hpos = xml.rindex("<", hpos_term)
        return unless hpos

        h_end_term = xml.index("/hyperlinks", hpos)
        h_end = h_end_term ? xml.index(">", h_end_term) : xml.size
        h_end ||= xml.size

        pos = hpos
        while pos < h_end
          hl_start_term = xml.index("hyperlink", pos)
          break unless hl_start_term && hl_start_term < h_end

          hl_start = xml.rindex("<", hl_start_term)
          break unless hl_start && hl_start < h_end

          hl_tag_end = xml.index(">", hl_start + 1)
          break unless hl_tag_end

          hl_tag = xml.byteslice(hl_start, hl_tag_end - hl_start)
          ref = tag_attr(hl_tag, ' ref="')
          rid = tag_attr(hl_tag, ' r:id="') || tag_attr(hl_tag, ' id="')
          display = tag_attr(hl_tag, ' display="')
          tooltip = tag_attr(hl_tag, ' tooltip="')
          location = tag_attr(hl_tag, ' location="')

          if ref
            block.call(Event.new(
                         type: :hyperlink,
                         args: [ref, rid, display, tooltip, location],
                         source: { part: part_name, cell: ref }
                       ))
          end
          pos = hl_tag_end + 1
        end
      end

      # ---- Fast string-scanning event parser (byte-level) ----
      #
      # SECURITY NOTE: This custom XML parsing approach is secure by design against
      # typical XML vulnerabilities:
      # - XXE (XML External Entity Expansion): It does not interpret DTDs or expand
      #   arbitrary entities. It purely scans for literal `<row>` and `<c>` tags.
      # - ReDoS: It uses O(1) byte indexing and bounded `String#index` searches
      #   rather than unbounded or backtracking regular expressions.

      def self.fast_scan_events(xml_src, shared_strings, part_name, &block)
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

          row_source = { part: part_name, row: row_index }
          block.call(Event.new(
                       type: :row_start,
                       args: [row_index, attrs],
                       source: row_source
                     ))

          fast_scan_cells_events(xml, tag_end + 1, row_end, shared_strings, row_source, &block)

          block.call(Event.new(
                       type: :row_end,
                       args: [],
                       source: row_source
                     ))

          pos = row_end + 6
        end
      end

      private_class_method :fast_scan_events

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

      def self.fast_scan_cells_events(xml, from, to, shared_strings, row_source, &block)
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
            cell_source = row_source.dup
            cell_source[:cell] = ref
            block.call(Event.new(
                         type: :cell,
                         args: [ref, type, style_index, nil, nil],
                         source: cell_source
                       ))
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

          val_to_use = inline_str || value
          cell_source = row_source.dup
          cell_source[:cell] = ref
          block.call(Event.new(
                       type: :cell,
                       args: [ref, type, style_index, val_to_use, formula],
                       source: cell_source
                     ))

          pos = c_end + 4
        end
      end

      private_class_method :fast_scan_cells_events

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
            raw.to_i
          end
        end
      end

      private_class_method :resolve_fast_value

      XML_ENTITIES = { "&amp;" => "&", "&lt;" => "<", "&gt;" => ">", "&quot;" => '"', "&apos;" => "'" }.freeze

      # SECURITY NOTE: Decodes only standard predefined XML entities in a single pass.
      # This completely prevents "Billion Laughs" attacks (exponential entity expansion)
      # because it avoids recursive expansion and ignores custom entities entirely.
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

      def self.fast_scan_rows_direct(xml_src, shared_strings, part_name, &block)
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

          row_index = 0
          has_custom_attrs = false
          ri = row_start + 4
          while ri < tag_end
            rb = xml.getbyte(ri)
            if rb == 114 && xml.getbyte(ri + 1) == 61 && xml.getbyte(ri + 2) == 34 # r="
              ri += 3
              while ri < tag_end
                cb = xml.getbyte(ri)
                break unless cb.between?(48, 57)

                row_index = (row_index * 10) + (cb - 48)
                ri += 1

              end
              row_index -= 1
            elsif [104, 99, 111].include?(rb) # 'h', 'c', 'o' for ht, customHeight, hidden, outlineLevel
              has_custom_attrs = true
              ri += 1
            else
              ri += 1
            end
          end

          attrs = if has_custom_attrs
                    row_tag = xml.byteslice(row_start, tag_end - row_start)
                    extract_row_attrs(row_tag)
                  else
                    EMPTY_HASH
                  end

          row_end = xml.index("</row>", tag_end + 1)
          break unless row_end

          { part: part_name, row: row_index }
          row_obj = StreamRow.new(
            index: row_index,
            xml_bytes: xml,
            from: tag_end + 1,
            to: row_end,
            shared_strings: shared_strings,
            height: attrs[:height],
            hidden: attrs[:hidden] || false,
            custom_height: attrs[:custom_height] || false,
            outline_level: attrs[:outline_level]
          )
          block.call(row_obj)

          pos = row_end + 6
        end
      end

      private_class_method :fast_scan_rows_direct

      def self.fast_parse_cells_direct(xml, from, to, shared_strings, row_source)
        cells = []
        fast_scan_cells_direct(xml, from, to, shared_strings, row_source) { |c| cells << c }
        cells
      end

      def self.fast_scan_cells_direct(xml, from, to, shared_strings, row_source, &block)
        pos = from
        cell_count = 0

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

          # Zero-allocation attribute scan directly within the tag
          type = nil
          style_index = nil
          col_idx = nil
          row_idx = nil

          ai = c_start + 2
          while ai < c_tag_end
            b = xml.getbyte(ai)
            if [32, 9, 10, 13].include?(b)
              ai += 1
              next
            end

            if b == 114 && xml.getbyte(ai + 1) == 61 && xml.getbyte(ai + 2) == 34 # r="
              ai += 3
              col_idx = 0
              while ai < c_tag_end
                cb = xml.getbyte(ai)
                if cb.between?(65, 90)
                  col_idx = (col_idx * 26) + (cb - 64)
                  ai += 1
                elsif cb.between?(97, 122)
                  col_idx = (col_idx * 26) + (cb - 96)
                  ai += 1
                else
                  break
                end
              end
              col_idx -= 1
              row_idx = 0
              while ai < c_tag_end
                cb = xml.getbyte(ai)
                break unless cb.between?(48, 57)

                row_idx = (row_idx * 10) + (cb - 48)
                ai += 1
              end
              row_idx -= 1
              ai += 1 if xml.getbyte(ai) == 34
            elsif b == 116 && xml.getbyte(ai + 1) == 61 && xml.getbyte(ai + 2) == 34 # t="
              ai += 3
              type_b = xml.getbyte(ai)
              if type_b == 115 && xml.getbyte(ai + 1) == 34 # "s"
                type = "s"
                ai += 2
              elsif type_b == 98 && xml.getbyte(ai + 1) == 34 # "b"
                type = "b"
                ai += 2
              elsif type_b == 101 && xml.getbyte(ai + 1) == 34 # "e"
                type = "e"
                ai += 2
              else
                t_end = xml.index('"', ai)
                type = xml.byteslice(ai, t_end - ai) if t_end
                ai = t_end ? t_end + 1 : c_tag_end
              end
            elsif b == 115 && xml.getbyte(ai + 1) == 61 && xml.getbyte(ai + 2) == 34 # s="
              ai += 3
              style_index = 0
              while ai < c_tag_end
                cb = xml.getbyte(ai)
                break unless cb.between?(48, 57)

                style_index = (style_index * 10) + (cb - 48)
                ai += 1
              end
              ai += 1 if xml.getbyte(ai) == 34
            else
              ai += 1
            end
          end

          # Self-closing <c ... />
          if xml.getbyte(c_tag_end - 1) == 47
            cell_obj = Elements::Cell.new(
              row_index: row_idx || row_source[:row],
              column_index: col_idx || cell_count,
              value: nil,
              style_index: style_index,
              errors: Elements::EMPTY_ERRORS
            )
            cell_count += 1
            block.call(cell_obj)
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
                  if type == "s"
                    # Zero-allocation integer parse for SST index
                    s_idx = 0
                    v_i = v_val_start
                    while v_i < v_end
                      s_idx = (s_idx * 10) + (xml.getbyte(v_i) - 48)
                      v_i += 1
                    end
                    value = shared_strings[s_idx] || ""
                  elsif type == "b"
                    value = xml.getbyte(v_val_start) == 49 # '1'
                  else
                    raw_value = xml.byteslice(v_val_start, v_end - v_val_start)
                    value = resolve_fast_value(raw_value, type, shared_strings)
                  end
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

          val_to_use = inline_str || value
          cell_obj = Elements::Cell.new(
            row_index: row_idx || row_source[:row],
            column_index: col_idx || cell_count,
            value: val_to_use,
            formula: formula,
            style_index: style_index,
            errors: Elements::EMPTY_ERRORS
          )
          cell_count += 1
          block.call(cell_obj)

          pos = c_end + 4
        end
      end

      private_class_method :fast_parse_cells_direct
    end
  end
end
