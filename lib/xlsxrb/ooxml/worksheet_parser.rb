# frozen_string_literal: true

# rbs_inline: enabled

require_relative "xml_parser"
require_relative "xml_builder"

module Xlsxrb
  module Ooxml
    # Streaming parser for xl/worksheets/sheetN.xml.
    # Uses a fast string-scanning approach for row/cell extraction,
    # falling back to REXML SAX only for column definitions and unmapped data.
    class WorksheetParser
      EMPTY_ARRAY = [].freeze
      EMPTY_HASH = {}.freeze
      TAG_DELIM_BYTES = [32, 62, 9, 10, 13, 47].freeze
      ROW_ATTR_START_BYTES = [104, 99, 111, 115].freeze

      # Parses all rows from a worksheet XML string.
      # Returns an Array of raw row hashes:
      #   { index:, cells: [{ ref:, type:, style_index:, value:, formula: }], attrs:, unmapped: }
      def self.parse(xml_string, shared_strings: [], styles: nil)
        return [] if xml_string.nil? || xml_string.empty?

        rows = []
        each_row(xml_string, shared_strings: shared_strings, styles: styles) { |row| rows << row }
        rows
      end

      # Streaming parse: yields one raw row hash at a time.
      def self.each_row(xml_source, shared_strings: [], part_name: "xl/worksheets/sheet1.xml", styles: nil, date1904: false, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, &block)
        return enum_for(:each_row, xml_source, shared_strings: shared_strings, part_name: part_name, styles: styles, date1904: date1904, trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells) unless block
        return if xml_source.nil? || (xml_source.respond_to?(:empty?) && xml_source.empty?)

        row_consumer = if trim_empty_rows
                         pending_empty = []
                         lambda do |row|
                           if row.empty?
                             pending_empty << row
                           else
                             pending_empty.each(&block)
                             pending_empty.clear
                             block.call(row)
                           end
                         end
                       else
                         block
                       end

        if pad_empty_rows
          current_expected_idx = 0
          inner_consumer = row_consumer
          row_consumer = lambda do |row|
            while current_expected_idx < row.index
              empty_row = StreamRow.fast_create(
                current_expected_idx, "".b, 0, 0, shared_strings, "", nil, false, false, nil, nil, styles, date1904, false, pad_empty_cells
              )
              inner_consumer.call(empty_row)
              current_expected_idx += 1
            end
            inner_consumer.call(row)
            current_expected_idx = row.index + 1
          end
        end

        if xml_source.is_a?(String)
          fast_scan_rows_direct(xml_source, shared_strings, part_name, styles, date1904, pad_empty_cells, &row_consumer)
        else
          scan_rows_stream(xml_source, shared_strings, part_name, styles, date1904, pad_empty_cells, &row_consumer)
        end
      end

      # Parses column definitions (<cols>) from a worksheet.
      def self.parse_columns(xml_string, _part_name = "xl/worksheets/sheet1.xml")
        return [] if xml_string.nil? || xml_string.empty?
        return [] unless xml_string.include?("<cols")

        listener = ColumnsListener.new
        XmlParser.parse(xml_string, listener)
        listener.columns
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
        parse_hyperlinks(xml_string).each do |hl|
          block.call(Event.new(
                       type: :hyperlink,
                       args: [hl[:ref], hl[:rid], hl[:display], hl[:tooltip], hl[:location]],
                       source: { part: part_name, cell: hl[:ref] }
                     ))
        end
      end

      # Parses all hyperlink elements from worksheet XML.
      #
      # @param xml_string [String, nil]
      # @return [Array<Hash[Symbol, untyped]>]
      # @api public
      #: (String?) -> Array[Hash[Symbol, untyped]]
      def self.parse_hyperlinks(xml_string)
        return [] if xml_string.nil? || xml_string.empty?
        return [] unless xml_string.include?("hyperlink")

        xml = xml_string.b
        hpos_term = xml.index("hyperlinks")
        return [] unless hpos_term

        hpos = xml.rindex("<", hpos_term)
        return [] unless hpos

        h_end_term = xml.index("/hyperlinks", hpos)
        h_end = h_end_term ? xml.index(">", h_end_term) : xml.size
        h_end ||= xml.size

        links = []
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
          rid = tag_attr(hl_tag, ' r:id="') || tag_attr(hl_tag, ' id="') || hl_tag[/(?:[a-zA-Z0-9_]+:)?id="([^"]+)"/, 1]
          display = tag_attr(hl_tag, ' display="')
          tooltip = tag_attr(hl_tag, ' tooltip="')
          location = tag_attr(hl_tag, ' location="')

          if ref
            entry = { ref: ref }
            entry[:rid] = rid if rid
            entry[:display] = display if display
            entry[:tooltip] = tooltip if tooltip
            entry[:location] = location if location
            links << entry
          end
          pos = hl_tag_end + 1
        end
        links
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
        last_row_index = -1

        while pos < sd_end
          row_start = xml.index(row_start_pattern, pos)
          break unless row_start && row_start < sd_end

          nb = xml.getbyte(row_start + row_start_len)
          unless TAG_DELIM_BYTES.include?(nb)
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
          r_val = tag_attr(row_tag, ' r="')
          row_index = r_val ? r_val.to_i - 1 : last_row_index + 1
          last_row_index = row_index

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

        if row_tag.include?('hidden="1"') || row_tag.include?('hidden="true"')
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:hidden] = true
        end

        if row_tag.include?('customHeight="1"') || row_tag.include?('customHeight="true"')
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:custom_height] = true
        end

        if row_tag.include?('collapsed="1"') || row_tag.include?('collapsed="true"')
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:collapsed] = true
        end

        ol_val = tag_attr(row_tag, ' outlineLevel="')
        if ol_val
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:outline_level] = ol_val.to_i
        end

        s_val = tag_attr(row_tag, ' s="')
        if s_val
          attrs = {} if attrs.equal?(EMPTY_HASH)
          attrs[:style_index] = s_val.to_i
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
          unless TAG_DELIM_BYTES.include?(nb)
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
        xml = xml.b
        result = +""
        pos = from
        while pos < to
          tag_start = xml.index("<", pos)
          break unless tag_start && tag_start < to

          # Check self-closing tag or closing tag
          next_byte = xml.getbyte(tag_start + 1)
          if next_byte == 47 # </...
            tag_end = xml.index(">", tag_start + 2)
            break unless tag_end

            pos = tag_end + 1
            next
          end

          tag_end = xml.index(">", tag_start + 1)
          break unless tag_end && tag_end <= to

          # Skip self-closing tags (<.../>)
          if xml.getbyte(tag_end - 1) == 47
            pos = tag_end + 1
            next
          end

          tag_header = xml.byteslice(tag_start + 1, tag_end - tag_start - 1)
          tag_name = tag_header.split(/[ \t\r\n]/, 2).first || ""
          tag_localname = tag_name.include?(":") ? tag_name.split(":", 2).last : tag_name

          if tag_localname == "rPh"
            # Skip phonetic readings
            close_rph = xml.index("</#{tag_name}>", tag_end + 1) || xml.index("</rPh>", tag_end + 1)
            break unless close_rph && close_rph < to

            close_end = xml.index(">", close_rph)
            break unless close_end

            pos = close_end + 1
            next
          elsif tag_localname == "t"
            close_t = xml.index("</#{tag_name}>", tag_end + 1) || xml.index("</t>", tag_end + 1)
            break unless close_t && close_t <= to

            close_end = xml.index(">", close_t)
            break unless close_end

            text_segment = xml.byteslice(tag_end + 1, close_t - tag_end - 1).force_encoding("UTF-8")
            result << normalize_text(text_segment)
            pos = close_end + 1
            next
          else
            pos = tag_end + 1
          end
        end
        result
      end

      private_class_method :extract_inline_text

      STRICT_NUMERIC_RE = /\A[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?\z/
      private_constant :STRICT_NUMERIC_RE

      # Resolves a cell value from its type attribute and parsed XML content (<v> or <is>).
      # Consolidates parsing across fast-scan direct, row values, events, and listener paths.
      #: (String? type, String? raw_value, String? is_xml, Array[String | Elements::RichText | untyped] shared_strings) -> untyped
      def self.resolve_cell_value(type, raw_value, is_xml, shared_strings)
        case type
        when "s"
          raw_value ? (shared_strings[raw_value.to_i] || "") : nil
        when "b"
          if raw_value
            v = raw_value.strip
            v == "1" || v.casecmp?("true")
          end
        when "inlineStr", "str", "e"
          if is_xml
            extract_inline_text(is_xml, 0, is_xml.bytesize)
          elsif raw_value
            normalize_text(raw_value)
          end
        when "d"
          raw_value && normalize_text(raw_value)
        when "n"
          raw_value ? numeric_cell_value(raw_value) : nil
        else
          if raw_value
            text = raw_value.strip
            return nil if text.empty?

            if text.match?(STRICT_NUMERIC_RE)
              Integer(text, 10, exception: false) ||
                Float(text, exception: false) ||
                normalize_text(raw_value)
            else
              normalize_text(raw_value)
            end
          end
        end
      end

      def self.resolve_fast_value(raw, type, shared_strings)
        resolve_cell_value(type, raw, nil, shared_strings)
      end

      private_class_method :resolve_fast_value

      # Converts the text of a numeric cell (<c t="n">) into an Integer or Float.
      # An empty <v></v> carries no value and yields nil. Text that is not an xsd:double
      # (e.g. "abc") cannot come from a conforming producer and raises ParseError
      # instead of silently becoming 0.
      #: (String raw) -> (Integer | Float)?
      def self.numeric_cell_value(raw)
        text = raw.strip
        return nil if text.empty?

        raise(ParseError, "invalid numeric cell value: #{text.inspect}") unless text.match?(STRICT_NUMERIC_RE)

        Integer(text, 10, exception: false) || Float(text, exception: false) ||
          raise(ParseError, "invalid numeric cell value: #{text.inspect}")
      end

      # SECURITY NOTE: Decodes only standard predefined XML entities in a single pass.
      # This completely prevents "Billion Laughs" attacks (exponential entity expansion)
      # because it avoids recursive expansion and ignores custom entities entirely.
      #: (String str) -> String
      def self.decode_xml_entities(str)
        XmlBuilder.unescape(str)
      end

      #: (String str) -> String
      def self.normalize_text(str)
        s = decode_xml_entities(str)
        s = s.gsub("\r\n", "\n").tr("\r", "\n") if s.include?("\r")
        s.encoding == Encoding::BINARY ? s.dup.force_encoding(Encoding::UTF_8) : s
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
              outline_level: attrs["outlineLevel"]&.to_i,
              style_index: attrs["style"]&.to_i
            }
            @columns << col
          end
        end

        def end_element(_uri, localname, _qname)
          @in_cols = false if localname == "cols"
        end

        def characters(_text); end
      end

      # rubocop:disable Style/OptionalBooleanParameter
      def self.fast_scan_rows_direct(xml_src, shared_strings, _part_name, styles = nil, date1904 = false, pad_empty_cells = false, &block)
        # rubocop:enable Style/OptionalBooleanParameter
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
        last_row_index = -1

        while pos < sd_end
          row_start = xml.index(row_start_pattern, pos)
          break unless row_start && row_start < sd_end

          nb = xml.getbyte(row_start + row_start_len)
          unless TAG_DELIM_BYTES.include?(nb)
            pos = row_start + row_start_len
            next
          end

          tag_end = xml.index(">", row_start + row_start_len)
          break unless tag_end

          if xml.getbyte(tag_end - 1) == 47
            pos = tag_end + 1
            next
          end

          row_index = nil
          has_custom_attrs = false
          ri = row_start + row_start_len
          while ri < tag_end
            rb = xml.getbyte(ri)
            if rb == 114 && xml.getbyte(ri + 1) == 61 && xml.getbyte(ri + 2) == 34 # r="
              ri += 3
              parsed_r = 0
              while ri < tag_end
                cb = xml.getbyte(ri)
                break unless cb.between?(48, 57)

                parsed_r = (parsed_r * 10) + (cb - 48)
                ri += 1
              end
              row_index = parsed_r - 1
            elsif ROW_ATTR_START_BYTES.include?(rb) # 'h', 'c', 'o', 's' for ht, customHeight, hidden, outlineLevel, s
              has_custom_attrs = true
              ri += 1
            else
              ri += 1
            end
          end

          row_index ||= last_row_index + 1
          last_row_index = row_index

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
            attrs[:outline_level],
            attrs[:style_index],
            styles,
            date1904,
            attrs[:collapsed] || false,
            pad_empty_cells
          )
          block.call(row_obj)

          pos = row_end + row_end_len
        end
      end

      private_class_method :fast_scan_rows_direct

      # rubocop:disable Style/OptionalBooleanParameter
      def self.scan_rows_stream(source, shared_strings, _part_name, styles = nil, date1904 = false, pad_empty_cells = false, &block)
        # rubocop:enable Style/OptionalBooleanParameter
        buffer = +""
        buffer.force_encoding(Encoding::BINARY)
        in_sheet_data = false
        prefix = ""
        row_start_pattern = "<row"
        row_start_len = 4
        row_end_tag = "</row>"
        row_end_len = 6
        sd_end_tag = "</sheetData>"
        last_row_index = -1

        push_chunk = lambda do |chunk|
          return if in_sheet_data == :closed
          return if chunk.nil? || chunk.empty?

          buffer << chunk.b

          unless in_sheet_data
            sd_term = buffer.index("sheetData")
            if sd_term.nil?
              buffer.slice!(0...(buffer.bytesize - 32)) if buffer.bytesize > 64
              return
            end

            sd_start = buffer.rindex("<", sd_term)
            if sd_start.nil?
              buffer.slice!(0...sd_term)
              return
            end

            sd_open_end = buffer.index(">", sd_term)
            return if sd_open_end.nil?

            if buffer.getbyte(sd_open_end - 1) == 47
              in_sheet_data = :closed
              buffer.clear
              return
            end

            prefix = buffer.byteslice(sd_start + 1, sd_term - (sd_start + 1))
            row_start_pattern = "<#{prefix}row"
            row_start_len = row_start_pattern.bytesize
            row_end_tag = "</#{prefix}row>"
            row_end_len = row_end_tag.bytesize
            sd_end_tag = "</#{prefix}sheetData>"

            buffer.slice!(0..sd_open_end)
            in_sheet_data = true
          end

          pos = 0
          last_consumed_pos = 0
          sd_end_pos = buffer.index(sd_end_tag)

          loop do
            row_start = buffer.index(row_start_pattern, pos)

            if sd_end_pos && (row_start.nil? || sd_end_pos < row_start)
              in_sheet_data = :closed
              buffer.clear
              return
            end

            break if row_start.nil?

            next_pos = row_start + row_start_len
            break if next_pos >= buffer.bytesize

            nb = buffer.getbyte(next_pos)
            unless TAG_DELIM_BYTES.include?(nb)
              pos = next_pos
              next
            end

            tag_end = buffer.index(">", next_pos)
            break if tag_end.nil?

            if buffer.getbyte(tag_end - 1) == 47
              pos = tag_end + 1
              last_consumed_pos = pos
              next
            end

            row_end = buffer.index(row_end_tag, tag_end + 1)
            if row_end.nil?
              raise ArgumentError, "Malformed XML: row size exceeds maximum buffer limit" if buffer.bytesize > 50 * 1024 * 1024

              break
            end

            row_index = nil
            has_custom_attrs = false
            ri = row_start + row_start_len
            while ri < tag_end
              rb = buffer.getbyte(ri)
              if rb == 114 && buffer.getbyte(ri + 1) == 61 && buffer.getbyte(ri + 2) == 34 # r="
                ri += 3
                parsed_r = 0
                while ri < tag_end
                  cb = buffer.getbyte(ri)
                  break unless cb.between?(48, 57)

                  parsed_r = (parsed_r * 10) + (cb - 48)
                  ri += 1
                end
                row_index = parsed_r - 1
              elsif ROW_ATTR_START_BYTES.include?(rb)
                has_custom_attrs = true
                ri += 1
              else
                ri += 1
              end
            end

            row_index ||= last_row_index + 1
            last_row_index = row_index

            attrs = if has_custom_attrs
                      row_tag = buffer.byteslice(row_start, tag_end - row_start)
                      extract_row_attrs(row_tag)
                    else
                      EMPTY_HASH
                    end

            cells_xml = buffer.byteslice(tag_end + 1, row_end - (tag_end + 1))
            row_obj = StreamRow.fast_create(
              row_index,
              cells_xml,
              0,
              cells_xml.bytesize,
              shared_strings,
              prefix,
              attrs[:height],
              attrs[:hidden] || false,
              attrs[:custom_height] || false,
              attrs[:outline_level],
              attrs[:style_index],
              styles,
              date1904,
              attrs[:collapsed] || false,
              pad_empty_cells
            )
            block.call(row_obj)

            pos = row_end + row_end_len
            last_consumed_pos = pos
          end

          buffer.slice!(0...last_consumed_pos) if last_consumed_pos.positive?
        end

        catch(:done) do
          if source.respond_to?(:each_chunk)
            source.each_chunk(65_536) do |chunk|
              push_chunk.call(chunk)
              throw :done if in_sheet_data == :closed
            end
          elsif source.respond_to?(:read)
            while (chunk = source.read(65_536))
              break if chunk.empty?

              push_chunk.call(chunk)
              break if in_sheet_data == :closed
            end
          elsif source.respond_to?(:call)
            source.call do |chunk|
              push_chunk.call(chunk)
              throw :done if in_sheet_data == :closed
            end
          elsif source.respond_to?(:each)
            source.each do |chunk|
              push_chunk.call(chunk)
              break if in_sheet_data == :closed
            end
          else
            raise ArgumentError, "Unsupported stream source: #{source.class}"
          end
        end
      end

      private_class_method :scan_rows_stream

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

      CELL_FAST_RE = %r{<(?:[a-zA-Z0-9_]+:)?c\b([^>]*(?<!/))>\s*(?:<(?:[a-zA-Z0-9_]+:)?f\b[^>]*?(?:/>|>([^<]*)</(?:[a-zA-Z0-9_]+:)?f>)\s*)?(?:<(?:[a-zA-Z0-9_]+:)?v>([^<]*)</(?:[a-zA-Z0-9_]+:)?v>\s*)?(?:<(?:[a-zA-Z0-9_]+:)?is>(.*?)</(?:[a-zA-Z0-9_]+:)?is>\s*)?.*?</(?:[a-zA-Z0-9_]+:)?c>|<(?:[a-zA-Z0-9_]+:)?c\b([^>]*)/?>}m
      CELL_GENERIC_RE = %r{<(?:[a-zA-Z0-9_]+:)?c\b([^>]*?)(?:>\s*(?:<(?:[a-zA-Z0-9_]+:)?f\b[^>]*?(?:/>|>([^<]*)</(?:[a-zA-Z0-9_]+:)?f>)\s*)?(?:<(?:[a-zA-Z0-9_]+:)?v>([^<]*)</(?:[a-zA-Z0-9_]+:)?v>\s*)?(?:<(?:[a-zA-Z0-9_]+:)?is>(.*?)</(?:[a-zA-Z0-9_]+:)?is>\s*)?.*?</(?:[a-zA-Z0-9_]+:)?c>|/>)}m
      ATTR_R = /\br="([A-Za-z0-9]+)"/
      ATTR_T = /\bt="([a-zA-Z]+)"/
      ATTR_S = /\bs="(\d+)"/

      def self.fast_parse_cells_direct(xml, from, to, shared_strings, row_source, prefix = "")
        cells = []
        fast_scan_cells_direct(xml, from, to, shared_strings, row_source, prefix) { |c| cells << c }
        cells
      end

      # rubocop:disable Style/OptionalBooleanParameter
      def self.fast_scan_cells_direct(xml, from, to, shared_strings, row_source, _prefix = "", styles = nil, date1904 = false, &block)
        # rubocop:enable Style/OptionalBooleanParameter
        chunk = xml.byteslice(from, to - from)
        row_idx = row_source.is_a?(Hash) ? row_source[:row] : row_source
        col_idx = 0
        matched_any = false

        chunk.scan(CELL_FAST_RE) do |attrs_str, f, v, is, self_attrs|
          matched_any = true
          raw_attrs = attrs_str || self_attrs || ""
          r = raw_attrs[ATTR_R, 1]
          s = raw_attrs.include?('s="') ? raw_attrs[ATTR_S, 1] : nil
          t = raw_attrs.include?('t="') ? raw_attrs[ATTR_T, 1] : nil
          style_idx = s&.to_i
          ref = r
          c_idx = if ref
                    expected_letter = COL_LETTERS[col_idx]
                    if expected_letter && ref.start_with?(expected_letter) && ref.getbyte(expected_letter.bytesize)&.between?(48, 57)
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
          v = v.strip if v
          f = f.strip if f

          val = resolve_cell_value(t, v, is, shared_strings)

          raw_val = if is
                      extract_inline_text(is, 0, is.bytesize)
                    elsif v
                      normalize_text(v)
                    end

          formula_expr = f ? normalize_text(f) : nil
          formula_expr = nil if formula_expr && formula_expr.empty?
          fmt_code = NumberFormatter.format_code_for(style_idx, styles) if styles && style_idx
          cell = Elements::Cell.fast_create(row_idx, c_idx, val, style_idx, formula_expr, raw_val, fmt_code, date1904)
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
                    if expected_letter && r.start_with?(expected_letter) && r.getbyte(expected_letter.bytesize)&.between?(48, 57)
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
          v = v.strip if v
          f = f.strip if f

          val = resolve_cell_value(t, v, is, shared_strings)

          raw_val = if is
                      extract_inline_text(is, 0, is.bytesize)
                    elsif v
                      normalize_text(v)
                    end

          style_idx = s&.to_i
          formula_expr = f ? normalize_text(f) : nil
          formula_expr = nil if formula_expr && formula_expr.empty?
          fmt_code = NumberFormatter.format_code_for(style_idx, styles) if styles && style_idx
          cell = Elements::Cell.fast_create(row_idx, c_idx, val, style_idx, formula_expr, raw_val, fmt_code, date1904)
          block.call(cell)
        end
      end

      # rubocop:disable Style/OptionalBooleanParameter
      #: (String xml, Integer from, Integer to, Array[String | Elements::RichText | untyped] shared_strings, ?untyped styles, ?bool date1904, ?type_cast: bool) -> Array[untyped]
      def self.fast_scan_row_values(xml, from, to, shared_strings, styles = nil, date1904 = false, type_cast: false)
        # rubocop:enable Style/OptionalBooleanParameter
        chunk = xml.byteslice(from, to - from)
        return [] if chunk.nil? || chunk.empty?

        col_idx = 0
        matched_any = false
        row_values = []

        chunk.scan(CELL_FAST_RE) do |attrs_str, _f, v, is, self_attrs|
          matched_any = true
          raw_attrs = attrs_str || self_attrs || ""
          r = raw_attrs[ATTR_R, 1]
          s_val = raw_attrs.include?('s="') ? raw_attrs[ATTR_S, 1] : nil
          t = raw_attrs.include?('t="') ? raw_attrs[ATTR_T, 1] : nil
          ref = r
          c_idx = if ref
                    expected_letter = COL_LETTERS[col_idx]
                    if expected_letter && ref.start_with?(expected_letter) && ref.getbyte(expected_letter.bytesize)&.between?(48, 57)
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
          v = v.strip if v

          val = resolve_cell_value(t, v, is, shared_strings)

          if type_cast && val.is_a?(Numeric) && s_val && styles
            fmt_type = NumberFormatter.format_type_for(s_val.to_i, styles)
            begin
              case fmt_type
              when :date
                val = Utils.serial_to_date(val, date1904: date1904)
              when :datetime, :time
                val = Utils.serial_to_datetime(val, date1904: date1904)
              end
            rescue StandardError
              # Keep numeric value on conversion failure
            end
          end

          row_values[c_idx] = val
        end

        return row_values if matched_any || (!chunk.include?("<c") && !chunk.include?(":c"))

        # Fallback for non-standard attribute ordering
        chunk.scan(CELL_GENERIC_RE) do |attrs, _f, v, is|
          r = attrs[ATTR_R, 1]
          t = attrs[ATTR_T, 1]
          s = attrs[ATTR_S, 1]

          c_idx = if r
                    expected_letter = COL_LETTERS[col_idx]
                    if expected_letter && r.start_with?(expected_letter) && r.getbyte(expected_letter.bytesize)&.between?(48, 57)
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
          v = v.strip if v

          val = resolve_cell_value(t, v, is, shared_strings)

          if type_cast && val.is_a?(Numeric) && s && styles
            fmt_type = NumberFormatter.format_type_for(s.to_i, styles)
            begin
              case fmt_type
              when :date
                val = Utils.serial_to_date(val, date1904: date1904)
              when :datetime, :time
                val = Utils.serial_to_datetime(val, date1904: date1904)
              end
            rescue StandardError
              # Keep numeric value on conversion failure
            end
          end

          row_values[c_idx] = val
        end

        row_values
      end

      private_class_method :fast_parse_cells_direct
    end
  end
end
