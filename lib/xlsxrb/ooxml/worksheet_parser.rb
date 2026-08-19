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

        sd_term = xml.index("sheetData")
        return unless sd_term

        sd_start = xml.rindex("<", sd_term)
        return unless sd_start

        prefix = xml.byteslice(sd_start + 1, sd_term - (sd_start + 1))

        sd_open_end = xml.index(">", sd_start)
        return unless sd_open_end

        return if xml.getbyte(sd_open_end - 1) == 47 # self-closing <sheetData/>

        sd_end_tag = "</#{prefix}sheetData>"
        sd_end = xml.index(sd_end_tag, sd_open_end)
        return unless sd_end

        row_start_pattern = "<#{prefix}row"
        row_start_len = row_start_pattern.bytesize
        row_end_tag = "</#{prefix}row>"
        row_end_len = row_end_tag.bytesize

        pos = sd_open_end + 1

        while pos < sd_end
          row_start = xml.index(row_start_pattern, pos)
          break unless row_start && row_start < sd_end

          nb = xml.getbyte(row_start + row_start_len)
          unless [32, 62, 9, 10, 13, 47].include?(nb)
            pos = row_start + row_start_len
            next
          end

          tag_end = xml.index(">", row_start + row_start_len)
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

          row_end = xml.index(row_end_tag, tag_end + 1)
          break unless row_end

          row_source = { part: part_name, row: row_index }
          block.call(Event.new(
                       type: :row_start,
                       args: [row_index, attrs],
                       source: row_source
                     ))

          fast_scan_cells_events(xml, tag_end + 1, row_end, shared_strings, row_source, prefix, &block)

          block.call(Event.new(
                       type: :row_end,
                       args: [],
                       source: row_source
                     ))

          pos = row_end + row_end_len
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

      def self.fast_scan_cells_events(xml, from, to, shared_strings, row_source, prefix = "", &block)
        pos = from
        c_start_pattern = "<#{prefix}c"
        c_start_len = c_start_pattern.bytesize
        c_end_tag = "</#{prefix}c>"
        c_end_len = c_end_tag.bytesize

        v_tag_start_str = "<#{prefix}v>"
        v_tag_start_len = v_tag_start_str.bytesize
        v_tag_end_str = "</#{prefix}v>"
        v_tag_end_len = v_tag_end_str.bytesize

        is_tag_start_str = "<#{prefix}is>"
        is_tag_start_len = is_tag_start_str.bytesize
        is_tag_end_str = "</#{prefix}is>"
        is_tag_end_len = is_tag_end_str.bytesize

        f_tag_start_str = "<#{prefix}f"
        f_tag_end_str = "</#{prefix}f>"
        f_tag_end_len = f_tag_end_str.bytesize

        while pos < to
          c_start = xml.index(c_start_pattern, pos)
          break unless c_start && c_start < to

          nb = xml.getbyte(c_start + c_start_len)
          unless [32, 62, 9, 10, 13, 47].include?(nb)
            pos = c_start + c_start_len
            next
          end

          c_tag_end = xml.index(">", c_start + c_start_len)
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

          c_end = xml.index(c_end_tag, c_tag_end + 1)
          break unless c_end

          # Parse cell content sequentially (bounded to c_end - avoid unbounded scans)
          value = nil
          formula = nil
          inline_str = nil
          cpos = c_tag_end + 1
          while cpos < c_end
            tag_pos = xml.index("<", cpos)
            break unless tag_pos && tag_pos < c_end

            if xml.byteslice(tag_pos, v_tag_start_len) == v_tag_start_str
              v_val_start = tag_pos + v_tag_start_len
              v_end = xml.index(v_tag_end_str, v_val_start)
              if v_end
                raw_value = xml.byteslice(v_val_start, v_end - v_val_start)
                value = resolve_fast_value(raw_value, type, shared_strings)
                cpos = v_end + v_tag_end_len
              else
                cpos = tag_pos + v_tag_start_len
              end
            elsif xml.byteslice(tag_pos, f_tag_start_str.bytesize) == f_tag_start_str
              f_tag_end = xml.index(">", tag_pos + f_tag_start_str.bytesize)
              if f_tag_end && f_tag_end < c_end
                if xml.getbyte(f_tag_end - 1) == 47 # self-closing <f ... />
                  cpos = f_tag_end + 1
                else
                  f_end = xml.index(f_tag_end_str, f_tag_end + 1)
                  if f_end && f_end <= c_end
                    formula = xml.byteslice(f_tag_end + 1, f_end - f_tag_end - 1).force_encoding("UTF-8")
                    formula = decode_xml_entities(formula) if formula.include?("&")
                    cpos = f_end + f_tag_end_len
                  else
                    cpos = f_tag_end + 1
                  end
                end
              else
                cpos = tag_pos + 2
              end
            elsif xml.byteslice(tag_pos, is_tag_start_len) == is_tag_start_str
              is_end = xml.index(is_tag_end_str, tag_pos + is_tag_start_len)
              if is_end && is_end <= c_end
                inline_str = extract_inline_text(xml, tag_pos + is_tag_start_len, is_end)
                cpos = is_end + is_tag_end_len
              else
                cpos = tag_pos + is_tag_start_len
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

          pos = c_end + c_end_len
        end
      end

      private_class_method :fast_scan_cells_events

      def self.extract_inline_text(xml, from, to)
        result = +""
        pos = from
        while pos < to
          t_start = xml.index("<", pos)
          break unless t_start && t_start < to

          t_term = xml.index("t", t_start)
          break unless t_term && t_term < to

          t_tag_end = xml.index(">", t_term)
          break unless t_tag_end && t_tag_end < to

          if xml.getbyte(t_tag_end - 1) == 47 # <t ... />
            pos = t_tag_end + 1
            next
          end

          t_close = xml.index("</", t_tag_end + 1)
          break unless t_close && t_close <= to

          t_end = xml.index(">", t_close + 2)
          break unless t_end && t_end <= to

          text_segment = xml.byteslice(t_tag_end + 1, t_close - t_tag_end - 1).force_encoding("UTF-8")
          text_segment = decode_xml_entities(text_segment) if text_segment.include?("&")
          result << text_segment

          pos = t_end + 1
        end
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

      def self.fast_scan_rows_direct(xml_src, shared_strings, _part_name, &block)
        xml = xml_src.b # force ASCII-8BIT for O(1) byte indexing

        sd_term = xml.index("sheetData")
        return unless sd_term

        sd_start = xml.rindex("<", sd_term)
        return unless sd_start

        prefix = xml.byteslice(sd_start + 1, sd_term - (sd_start + 1))

        sd_open_end = xml.index(">", sd_start)
        return unless sd_open_end
        return if xml.getbyte(sd_open_end - 1) == 47 # self-closing <sheetData/>

        sd_end_tag = "</#{prefix}sheetData>"
        sd_end = xml.index(sd_end_tag, sd_open_end)
        return unless sd_end

        row_start_pattern = "<#{prefix}row"
        row_start_len = row_start_pattern.bytesize
        row_end_tag = "</#{prefix}row>"
        row_end_len = row_end_tag.bytesize

        pos = sd_open_end + 1

        while pos < sd_end
          row_start = xml.index(row_start_pattern, pos)
          break unless row_start && row_start < sd_end

          nb = xml.getbyte(row_start + row_start_len)
          unless [32, 62, 9, 10, 13, 47].include?(nb)
            pos = row_start + row_start_len
            next
          end

          tag_end = xml.index(">", row_start + row_start_len)
          break unless tag_end

          if xml.getbyte(tag_end - 1) == 47
            pos = tag_end + 1
            next
          end

          row_index = 0
          has_custom_attrs = false
          ri = row_start + row_start_len
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

          row_end = xml.index(row_end_tag, tag_end + 1)
          break unless row_end

          row_obj = StreamRow.fast_create(
            row_index,
            xml,
            tag_end + 1,
            row_end,
            shared_strings,
            prefix,
            attrs[:height],
            attrs[:hidden] || false,
            attrs[:custom_height] || false,
            attrs[:outline_level]
          )
          block.call(row_obj)

          pos = row_end + row_end_len
        end
      end

      private_class_method :fast_scan_rows_direct

      COL_LETTERS = (0...16_384).map do |index|
        result = +""
        i = index
        loop do
          result.prepend(("A".ord + (i % 26)).chr)
          i = (i / 26) - 1
          break if i.negative?
        end
        result.freeze
      end.freeze

      COL_MAP = (0...16_384).to_h do |idx|
        [COL_LETTERS[idx], idx]
      end.freeze

      CELL_FAST_RE = %r{<(?:[a-zA-Z0-9_]+:)?c\s+r="([A-Za-z0-9]+)"(?:[^>]*?s="(\d+)")?(?:[^>]*?t="([a-zA-Z]+)")?[^>]*>(?:<(?:[a-zA-Z0-9_]+:)?f[^>]*>([^<]*)</(?:[a-zA-Z0-9_]+:)?f>)?(?:<(?:[a-zA-Z0-9_]+:)?v>([^<]*)</(?:[a-zA-Z0-9_]+:)?v>)?(?:<(?:[a-zA-Z0-9_]+:)?is>(.*?)</(?:[a-zA-Z0-9_]+:)?is>)?</(?:[a-zA-Z0-9_]+:)?c>|<(?:[a-zA-Z0-9_]+:)?c\s+r="([A-Za-z0-9]+)"(?:[^>]*?s="(\d+)")?[^>]*/>}
      CELL_GENERIC_RE = %r{<(?:[a-zA-Z0-9_]+:)?c\b([^>]*?)(?:>(?:<(?:[a-zA-Z0-9_]+:)?f[^>]*>([^<]*)</(?:[a-zA-Z0-9_]+:)?f>)?(?:<(?:[a-zA-Z0-9_]+:)?v>([^<]*)</(?:[a-zA-Z0-9_]+:)?v>)?(?:<(?:[a-zA-Z0-9_]+:)?is>(.*?)</(?:[a-zA-Z0-9_]+:)?is>)?</(?:[a-zA-Z0-9_]+:)?c>|/>)}m
      ATTR_R = /r="([A-Za-z0-9]+)"/
      ATTR_T = /t="([a-zA-Z]+)"/
      ATTR_S = /s="(\d+)"/

      def self.fast_parse_cells_direct(xml, from, to, shared_strings, row_source, prefix = "")
        cells = []
        fast_scan_cells_direct(xml, from, to, shared_strings, row_source, prefix) { |c| cells << c }
        cells
      end

      def self.fast_scan_cells_direct(xml, from, to, shared_strings, row_source, _prefix = "", &block)
        chunk = xml.byteslice(from, to - from)
        row_idx = row_source.is_a?(Hash) ? row_source[:row] : row_source
        col_idx = 0
        matched_any = false

        chunk.scan(CELL_FAST_RE) do |r, s, t, f, v, is, self_r, self_s|
          matched_any = true
          ref = r || self_r
          s_val = s || self_s
          style_idx = s_val&.to_i
          c_idx = if ref
                    expected_letter = COL_LETTERS[col_idx]
                    if expected_letter && ref.start_with?(expected_letter)
                      col_idx
                    else
                      c_len = 0
                      c_len += 1 while (b = ref.getbyte(c_len)) && (b.between?(65, 90) || b.between?(97, 122))
                      col_letters = ref.byteslice(0, c_len).upcase
                      COL_MAP[col_letters] || col_idx
                    end
                  else
                    col_idx
                  end
          col_idx = c_idx + 1

          val = if t == "s"
                  shared_strings[v.to_i] || ""
                elsif t == "b"
                  v == "1"
                elsif %w[inlineStr str e].include?(t)
                  if is
                    extract_inline_text(is, 0, is.bytesize)
                  elsif v
                    v.include?("&") ? decode_xml_entities(v) : v
                  else
                    ""
                  end
                elsif v
                  v.include?(".") ? v.to_f : v.to_i
                end

          formula_expr = f&.include?("&") ? decode_xml_entities(f) : f
          cell = Elements::Cell.fast_create(row_idx, c_idx, val, style_idx, formula_expr)
          block.call(cell)
        end

        return if matched_any || (!chunk.include?("<c") && !chunk.include?(":c"))

        # Fallback for non-standard attribute ordering
        chunk.scan(CELL_GENERIC_RE) do |attrs, f, v, is|
          r = attrs[ATTR_R, 1]
          t = attrs[ATTR_T, 1]
          s = attrs[ATTR_S, 1]

          c_idx = if r
                    expected_letter = COL_LETTERS[col_idx]
                    if expected_letter && r.start_with?(expected_letter)
                      col_idx
                    else
                      c_len = 0
                      c_len += 1 while (b = r.getbyte(c_len)) && (b.between?(65, 90) || b.between?(97, 122))
                      col_letters = r.byteslice(0, c_len).upcase
                      COL_MAP[col_letters] || col_idx
                    end
                  else
                    col_idx
                  end
          col_idx = c_idx + 1

          val = if t == "s"
                  shared_strings[v.to_i] || ""
                elsif t == "b"
                  v == "1"
                elsif %w[inlineStr str e].include?(t)
                  if is
                    extract_inline_text(is, 0, is.bytesize)
                  elsif v
                    v.include?("&") ? decode_xml_entities(v) : v
                  else
                    ""
                  end
                elsif v
                  v.include?(".") ? v.to_f : v.to_i
                end

          style_idx = s&.to_i
          formula_expr = f&.include?("&") ? decode_xml_entities(f) : f
          cell = Elements::Cell.fast_create(row_idx, c_idx, val, style_idx, formula_expr)
          block.call(cell)
        end
      end

      private_class_method :fast_parse_cells_direct
    end
  end
end
