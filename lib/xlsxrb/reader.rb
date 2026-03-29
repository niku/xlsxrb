# frozen_string_literal: true

require "zlib"
require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  # Reads cells from an XLSX file.
  class Reader
    def initialize(filepath)
      @filepath = filepath
    end

    # Returns cells for the given sheet (by name or 0-based index).
    # Defaults to the first sheet. Numeric cells with date numFmt are converted to Date.
    def cells(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      shared_strings = load_shared_strings
      raw_cells = parse_worksheet_cells(worksheet_xml, shared_strings)

      # Resolve date-formatted cells.
      styles = load_styles
      return raw_cells if styles.empty?

      cell_style_map = parse_cell_style_indices(worksheet_xml)
      resolve_date_cells(raw_cells, cell_style_map, styles)
    end

    # Returns column widths as { "A" => 20.0, "B" => 15.5 } for the given sheet.
    def columns(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_columns(worksheet_xml)
    end

    # Returns row attributes as { 1 => { height: 25.0 }, 3 => { hidden: true } }.
    def row_attributes(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_row_attributes(worksheet_xml)
    end

    # Returns merged cell ranges as ["A1:B2", "C3:D4"].
    def merged_cells(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return [] if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_merge_cells(worksheet_xml)
    end

    # Returns hyperlinks as { "A1" => "https://example.com" }.
    def hyperlinks(sheet: nil)
      sheets = discover_sheets
      raise ArgumentError, "workbook has no sheets" if sheets.empty?

      target = resolve_sheet_target(sheets, sheet)
      raise ArgumentError, "sheet not found: #{sheet.inspect}" if target.nil?

      entry_path = if target.start_with?("/")
                     target.delete_prefix("/")
                   else
                     "xl/#{target}"
                   end

      worksheet_xml = extract_zip_entry(entry_path)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      # Parse hyperlink elements from worksheet.
      parser = REXML::Parsers::SAX2Parser.new(worksheet_xml)
      listener = HyperlinksListener.new
      parser.listen(listener)
      parser.parse

      # Parse rels to resolve rId -> URL.
      rels_path = entry_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
      rels_xml = extract_zip_entry(rels_path)
      rid_to_url = {}
      rid_to_url = parse_rels(rels_xml).transform_values { |v| v } if rels_xml && !rels_xml.empty?

      result = {}
      listener.links.each do |link|
        url = rid_to_url[link[:rid]]
        result[link[:ref]] = url if url
      end
      result
    end

    # Returns cell format codes as { "A1" => "0.00" } for cells with custom numFmt.
    def cell_formats(sheet: nil)
      # Load styles.
      styles = load_styles
      return {} if styles.empty?

      # Parse worksheet to get cell style indices.
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(worksheet_xml)
      listener = CellStyleListener.new
      parser.listen(listener)
      parser.parse

      result = {}
      listener.cell_style_indices.each do |cell_ref, xf_index|
        xf = styles[:cell_xfs][xf_index]
        next unless xf

        fmt_id = xf[:num_fmt_id]
        next unless fmt_id && fmt_id != 0

        format_code = styles[:num_fmts][fmt_id]
        result[cell_ref] = format_code if format_code
      end
      result
    end

    # Returns the autoFilter range string (e.g. "A1:B10") or nil.
    def auto_filter(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_auto_filter(worksheet_xml)
    end

    # Returns sheet-level properties (tabColor, outlinePr) for the given sheet.
    def sheet_properties(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return {} if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_properties(worksheet_xml)
    end

    # Returns the dimension ref string (e.g. "A1:B10") for the given sheet.
    def dimension(sheet: nil)
      worksheet_xml = load_worksheet_xml(sheet)
      return nil if worksheet_xml.nil? || worksheet_xml.empty?

      parse_worksheet_dimension(worksheet_xml)
    end

    # Returns core properties as a hash (e.g. { title: "...", creator: "..." }).
    def core_properties
      # Discover core properties path from _rels/.rels
      rels_xml = extract_zip_entry("_rels/.rels")
      return {} if rels_xml.nil? || rels_xml.empty?

      rels = parse_rels_with_types(rels_xml)
      core_rel = rels.find { |r| r[:type]&.end_with?("/metadata/core-properties") }
      return {} unless core_rel

      target = core_rel[:target]
      entry_path = target.start_with?("/") ? target.delete_prefix("/") : target
      xml = extract_zip_entry(entry_path)
      return {} if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CorePropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    # Returns app properties as a hash.
    def app_properties
      # Try standard path first, then discover via rels
      xml = extract_zip_entry("docProps/app.xml")
      if xml.nil? || xml.empty?
        rels_xml = extract_zip_entry("_rels/.rels")
        return {} if rels_xml.nil? || rels_xml.empty?

        rels = parse_rels_with_types(rels_xml)
        app_rel = rels.find { |r| r[:type]&.end_with?("/extended-properties") }
        return {} unless app_rel

        target = app_rel[:target]
        entry_path = target.start_with?("/") ? target.delete_prefix("/") : target
        xml = extract_zip_entry(entry_path)
      end
      return {} if xml.nil? || xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = AppPropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    # Returns workbook properties (e.g. { date1904: false, default_theme_version: 166925 }).
    def workbook_properties
      parse_workbook_metadata[:workbook_properties]
    end

    # Returns workbook view properties (e.g. { active_tab: 0 }).
    def workbook_views
      parse_workbook_metadata[:workbook_views]
    end

    # Returns calc properties (e.g. { calc_id: 191029 }).
    def calc_properties
      parse_workbook_metadata[:calc_properties]
    end

    # Returns defined names as an array of hashes.
    def defined_names
      parse_workbook_metadata[:defined_names]
    end

    # Returns sheet states as { "Sheet1" => :visible, "Hidden" => :hidden }.
    def sheet_states
      sheets = discover_sheets
      result = {}
      sheets.each do |s|
        state = case s[:state]
                when "hidden" then :hidden
                when "veryHidden" then :very_hidden
                else :visible
                end
        result[s[:name]] = state
      end
      result
    end

    # Returns ordered sheet names.
    def sheet_names
      discover_sheets.map { |s| s[:name] }
    end

    private

    def parse_workbook_metadata
      workbook_xml = extract_zip_entry("xl/workbook.xml")
      return { workbook_properties: {}, workbook_views: {}, calc_properties: {} } if workbook_xml.nil? || workbook_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(workbook_xml)
      listener = WorkbookListener.new
      parser.listen(listener)
      parser.parse
      {
        workbook_properties: listener.workbook_properties,
        workbook_views: listener.workbook_views,
        calc_properties: listener.calc_properties,
        defined_names: listener.defined_names
      }
    end

    def load_worksheet_xml(sheet)
      sheets = discover_sheets
      raise ArgumentError, "workbook has no sheets" if sheets.empty?

      target = resolve_sheet_target(sheets, sheet)
      raise ArgumentError, "sheet not found: #{sheet.inspect}" if target.nil?

      # Target may be absolute (/xl/worksheets/sheet1.xml) or relative (worksheets/sheet1.xml).
      entry_path = if target.start_with?("/")
                     target.delete_prefix("/")
                   else
                     "xl/#{target}"
                   end

      extract_zip_entry(entry_path)
    end

    def discover_sheets
      workbook_xml = extract_zip_entry("xl/workbook.xml")
      return [{ name: "Sheet1", rid: "rId1", target: "worksheets/sheet1.xml" }] if workbook_xml.nil? || workbook_xml.empty?

      rels_xml = extract_zip_entry("xl/_rels/workbook.xml.rels")
      rid_to_target = parse_rels(rels_xml)

      sheets = []
      parser = REXML::Parsers::SAX2Parser.new(workbook_xml)
      listener = WorkbookListener.new
      parser.listen(listener)
      parser.parse

      listener.sheets.each do |s|
        target = rid_to_target[s[:rid]]
        sheets << { name: s[:name], rid: s[:rid], target: target, state: s[:state] } if target
      end
      sheets
    end

    def parse_rels(rels_xml)
      return {} if rels_xml.nil? || rels_xml.empty?

      mapping = {}
      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = RelsListener.new
      parser.listen(listener)
      parser.parse
      listener.relationships.each { |r| mapping[r[:id]] = r[:target] }
      mapping
    end

    def parse_rels_with_types(rels_xml)
      return [] if rels_xml.nil? || rels_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(rels_xml)
      listener = RelsListener.new
      parser.listen(listener)
      parser.parse
      listener.relationships
    end

    def resolve_sheet_target(sheets, sheet)
      case sheet
      when nil
        sheets.first&.dig(:target)
      when Integer
        sheets[sheet]&.dig(:target)
      when String
        sheets.find { |s| s[:name] == sheet }&.dig(:target)
      else
        raise ArgumentError, "sheet must be a String name or Integer index"
      end
    end

    def load_shared_strings
      sst_xml = extract_zip_entry("xl/sharedStrings.xml")
      return [] if sst_xml.nil? || sst_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(sst_xml)
      listener = SharedStringsListener.new
      parser.listen(listener)
      parser.parse
      listener.strings
    end

    def load_styles
      styles_xml = extract_zip_entry("xl/styles.xml")
      return {} if styles_xml.nil? || styles_xml.empty?

      parser = REXML::Parsers::SAX2Parser.new(styles_xml)
      listener = StylesListener.new
      parser.listen(listener)
      parser.parse
      { num_fmts: listener.num_fmts, cell_xfs: listener.cell_xfs }
    end

    def parse_cell_style_indices(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = CellStyleListener.new
      parser.listen(listener)
      parser.parse
      listener.cell_style_indices
    end

    def resolve_date_cells(raw_cells, cell_style_map, styles)
      raw_cells.each do |cell_ref, value|
        next unless value.is_a?(Numeric)

        xf_index = cell_style_map[cell_ref]
        next unless xf_index

        xf = styles[:cell_xfs][xf_index]
        next unless xf

        fmt_id = xf[:num_fmt_id]
        next unless date_format?(fmt_id, styles[:num_fmts])

        raw_cells[cell_ref] = Xlsxrb.serial_to_date(value.to_i)
      end
      raw_cells
    end

    def date_format?(fmt_id, custom_num_fmts)
      return false unless fmt_id

      # Built-in date format IDs.
      return true if BUILTIN_DATE_FMT_IDS.include?(fmt_id)

      # Check custom format code for date-like patterns.
      code = custom_num_fmts[fmt_id]
      return false unless code

      date_pattern?(code)
    end

    def date_pattern?(code)
      # Strip quoted strings to avoid false matches.
      stripped = code.gsub(/"[^"]*"/, "").gsub(/\\[.]/, "")
      stripped.match?(/[ymdhsYMDHS]/)
    end

    def extract_zip_entry(entry_name)
      File.open(@filepath, "rb") do |file|
        loop do
          signature = file.read(4)
          break if signature.nil? || signature.bytesize < 4

          signature_value = signature.unpack1("V")
          break if [0x02014b50, 0x06054b50].include?(signature_value)

          raise Error, "invalid ZIP local header signature" unless signature_value == 0x04034b50

          header = file.read(26)
          raise Error, "truncated ZIP local header" if header.nil? || header.bytesize < 26

          _version, flags, compression_method, _mtime, _mdate, _crc32, compressed_size,
            _uncompressed_size, file_name_length, extra_field_length = header.unpack("v v v v v V V V v v")

          raise Error, "ZIP data descriptor is not supported" if flags.anybits?(0x0008)

          file_name = file.read(file_name_length)
          raise Error, "truncated ZIP file name" if file_name.nil? || file_name.bytesize < file_name_length

          file.read(extra_field_length)

          compressed_data = file.read(compressed_size)
          raise Error, "truncated ZIP entry data" if compressed_data.nil? || compressed_data.bytesize < compressed_size

          next unless file_name == entry_name

          case compression_method
          when 0
            return compressed_data
          when 8
            inflater = Zlib::Inflate.new(-Zlib::MAX_WBITS)
            begin
              return inflater.inflate(compressed_data)
            ensure
              inflater.close
            end
          else
            raise Error, "unsupported ZIP compression method: #{compression_method}"
          end
        end
      end

      nil
    end

    def parse_worksheet_cells(xml, shared_strings)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = WorksheetListener.new(shared_strings)
      parser.listen(listener)
      parser.parse
      listener.cells
    end

    def parse_worksheet_columns(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = ColumnsListener.new
      parser.listen(listener)
      parser.parse
      listener.raw_columns.transform_keys { |idx| column_index_to_letter(idx) }
    end

    def parse_worksheet_row_attributes(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = WorksheetListener.new([])
      parser.listen(listener)
      parser.parse
      listener.row_attributes
    end

    def parse_worksheet_merge_cells(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = MergeCellsListener.new
      parser.listen(listener)
      parser.parse
      listener.ranges
    end

    def parse_worksheet_auto_filter(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = AutoFilterListener.new
      parser.listen(listener)
      parser.parse
      listener.ref
    end

    def parse_worksheet_properties(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = SheetPropertiesListener.new
      parser.listen(listener)
      parser.parse
      listener.properties
    end

    def parse_worksheet_dimension(xml)
      parser = REXML::Parsers::SAX2Parser.new(xml)
      listener = DimensionListener.new
      parser.listen(listener)
      parser.parse
      listener.ref
    end

    def column_index_to_letter(index)
      result = +""
      while index.positive?
        index -= 1
        result.prepend(("A".ord + (index % 26)).chr)
        index /= 26
      end
      result
    end

    # SAX2 listener for parsing shared string table (xl/sharedStrings.xml).
    class SharedStringsListener
      include REXML::SAX2Listener

      attr_reader :strings

      def initialize
        @strings = []
        @inside_si = false
        @inside_t = false
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, _attributes)
        name = element_name(local_name, qname)

        case name
        when "si"
          @inside_si = true
          @text_buffer = +""
        when "t"
          @inside_t = @inside_si
        end
      end

      def characters(text)
        @text_buffer << text if @inside_t
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "t"
          @inside_t = false
        when "si"
          @strings << @text_buffer.dup
          @inside_si = false
          @text_buffer = +""
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing worksheet cells.
    class WorksheetListener
      include REXML::SAX2Listener

      attr_reader :cells, :row_attributes

      def initialize(shared_strings = [])
        @shared_strings = shared_strings
        @cells = {}
        @row_attributes = {}
        @current_cell_ref = nil
        @current_cell_type = nil
        @inside_value = false
        @inside_inline_text = false
        @inside_formula = false
        @value_buffer = +""
        @inline_text_buffer = +""
        @formula_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "row"
          parse_row_attributes(attributes)
        when "c"
          @current_cell_ref = attributes["r"]
          @current_cell_type = attributes["t"]
          @value_buffer = +""
          @inline_text_buffer = +""
          @formula_buffer = +""
        when "v"
          @inside_value = true
        when "f"
          @inside_formula = true
        when "t"
          @inside_inline_text = @current_cell_type == "inlineStr" && !@current_cell_ref.nil?
        end
      end

      def characters(text)
        @value_buffer << text if @inside_value
        @inline_text_buffer << text if @inside_inline_text
        @formula_buffer << text if @inside_formula
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)

        case name
        when "v"
          @inside_value = false
        when "f"
          @inside_formula = false
        when "t"
          @inside_inline_text = false
        when "c"
          store_cell_value
          @current_cell_ref = nil
          @current_cell_type = nil
          @value_buffer = +""
          @inline_text_buffer = +""
          @formula_buffer = +""
        end
      end

      private

      def parse_row_attributes(attributes)
        row_num = attributes["r"]&.to_i
        return unless row_num

        attrs = {}
        ht = attributes["ht"]
        attrs[:height] = ht.to_f if ht && attributes["customHeight"] == "1"
        attrs[:hidden] = true if attributes["hidden"] == "1"
        @row_attributes[row_num] = attrs unless attrs.empty?
      end

      def store_cell_value
        return if @current_cell_ref.nil?

        unless @formula_buffer.empty?
          cached = @value_buffer.empty? ? nil : @value_buffer.dup
          @cells[@current_cell_ref] = Formula.new(expression: @formula_buffer.dup, cached_value: cached)
          return
        end

        case @current_cell_type
        when "inlineStr"
          @cells[@current_cell_ref] = @inline_text_buffer.dup
        when "s"
          index = @value_buffer.to_i
          @cells[@current_cell_ref] = @shared_strings[index] || ""
        when "e"
          @cells[@current_cell_ref] = @value_buffer.dup
        when "b"
          @cells[@current_cell_ref] = @value_buffer.strip == "1"
        when nil, "", "n"
          return if @value_buffer.empty?

          raw = @value_buffer.dup
          @cells[@current_cell_ref] = numeric_value(raw)
        end
      end

      def numeric_value(raw)
        if raw.include?(".")
          raw.to_f
        else
          raw.to_i
        end
      end

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing workbook.xml to discover sheet names, rIds, and workbook-level properties.
    class WorkbookListener
      include REXML::SAX2Listener

      attr_reader :sheets, :workbook_properties, :workbook_views, :calc_properties, :defined_names

      def initialize
        @sheets = []
        @workbook_properties = {}
        @workbook_views = {}
        @calc_properties = {}
        @defined_names = []
        @inside_defined_name = false
        @current_dn_attrs = nil
        @dn_text_buffer = +""
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "sheet"
          @sheets << { name: attributes["name"], rid: attributes["r:id"], state: attributes["state"] }
        when "workbookPr"
          d1904 = attributes["date1904"]
          @workbook_properties[:date1904] = %w[1 true].include?(d1904) unless d1904.nil?
          dtv = attributes["defaultThemeVersion"]
          @workbook_properties[:default_theme_version] = dtv.to_i if dtv
        when "workbookView"
          at = attributes["activeTab"]
          @workbook_views[:active_tab] = at.to_i if at
          fs = attributes["firstSheet"]
          @workbook_views[:first_sheet] = fs.to_i if fs
        when "calcPr"
          ci = attributes["calcId"]
          @calc_properties[:calc_id] = ci.to_i if ci
          fcol = attributes["fullCalcOnLoad"]
          @calc_properties[:full_calc_on_load] = %w[1 true].include?(fcol) unless fcol.nil?
        when "definedName"
          @inside_defined_name = true
          @current_dn_attrs = {
            name: attributes["name"],
            hidden: %w[1 true].include?(attributes["hidden"])
          }
          lsi = attributes["localSheetId"]
          @current_dn_attrs[:local_sheet_id] = lsi.to_i if lsi
          @dn_text_buffer = +""
        end
      end

      def characters(text)
        @dn_text_buffer << text if @inside_defined_name
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless name == "definedName" && @inside_defined_name

        @current_dn_attrs[:value] = @dn_text_buffer.dup
        @defined_names << @current_dn_attrs
        @inside_defined_name = false
        @current_dn_attrs = nil
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing .rels files to map rId to Target.
    class RelsListener
      include REXML::SAX2Listener

      attr_reader :relationships

      def initialize
        @relationships = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "Relationship"

        @relationships << { id: attributes["Id"], target: attributes["Target"], type: attributes["Type"] }
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <cols><col> elements from worksheet XML.
    class ColumnsListener
      include REXML::SAX2Listener

      # Returns { column_index => width } hash (1-based indices).
      attr_reader :raw_columns

      def initialize
        @raw_columns = {}
      end

      def columns
        @raw_columns
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "col"

        min_val = attributes["min"]&.to_i
        max_val = attributes["max"]&.to_i
        width = attributes["width"]&.to_f
        return unless min_val && max_val && width

        (min_val..max_val).each { |i| @raw_columns[i] = width }
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <mergeCells><mergeCell> elements.
    class MergeCellsListener
      include REXML::SAX2Listener

      attr_reader :ranges

      def initialize
        @ranges = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "mergeCell"

        ref = attributes["ref"]
        @ranges << ref if ref
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <hyperlinks><hyperlink> elements.
    class HyperlinksListener
      include REXML::SAX2Listener

      attr_reader :links

      def initialize
        @links = []
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "hyperlink"

        ref = attributes["ref"]
        rid = attributes["r:id"]
        @links << { ref: ref, rid: rid } if ref && rid
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing styles.xml (numFmts + cellXfs).
    class StylesListener
      include REXML::SAX2Listener

      attr_reader :num_fmts, :cell_xfs

      def initialize
        @num_fmts = {} # { numFmtId => formatCode }
        @cell_xfs = [] # Array of { num_fmt_id: N }
        @inside_cell_xfs = false
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)

        case name
        when "numFmt"
          id = attributes["numFmtId"]&.to_i
          code = attributes["formatCode"]
          @num_fmts[id] = code if id && code
        when "cellXfs"
          @inside_cell_xfs = true
        when "xf"
          @cell_xfs << { num_fmt_id: attributes["numFmtId"]&.to_i } if @inside_cell_xfs
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        @inside_cell_xfs = false if name == "cellXfs"
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <autoFilter> element.
    class AutoFilterListener
      include REXML::SAX2Listener

      attr_reader :ref

      def initialize
        @ref = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "autoFilter"

        @ref = attributes["ref"]
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing docProps/app.xml.
    class AppPropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      def initialize
        @properties = {}
        @current_field = nil
        @text_buffer = +""
        @inside_vector = false
        @vector_items = []
        @heading_pairs = []
        @titles_of_parts = []
        @inside_heading_pairs = false
        @inside_titles_of_parts = false
        @inside_variant = false
      end

      def start_element(_uri, local_name, qname, _attributes)
        name = element_name(local_name, qname)
        case name
        when "Application", "AppVersion"
          @current_field = name
          @text_buffer = +""
        when "HeadingPairs"
          @inside_heading_pairs = true
        when "TitlesOfParts"
          @inside_titles_of_parts = true
        when "variant"
          @inside_variant = true
          @text_buffer = +""
        when "lpstr", "i4"
          @text_buffer = +""
        end
      end

      def characters(text)
        @text_buffer << text
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        case name
        when "Application"
          @properties[:application] = @text_buffer.dup
          @current_field = nil
        when "AppVersion"
          @properties[:app_version] = @text_buffer.dup
          @current_field = nil
        when "lpstr"
          if @inside_titles_of_parts
            @titles_of_parts << @text_buffer.dup
          elsif @inside_heading_pairs && @inside_variant
            @vector_items << @text_buffer.dup
          end
        when "i4"
          @vector_items << @text_buffer.to_i if @inside_heading_pairs && @inside_variant
        when "variant"
          @inside_variant = false
        when "HeadingPairs"
          # Convert flat array to pairs: [label, count, label, count, ...]
          @heading_pairs = @vector_items.each_slice(2).to_a
          @vector_items = []
          @inside_heading_pairs = false
          @properties[:heading_pairs] = @heading_pairs
        when "TitlesOfParts"
          @inside_titles_of_parts = false
          @properties[:titles_of_parts] = @titles_of_parts
        end
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing docProps/core.xml.
    class CorePropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      FIELD_MAP = {
        "title" => :title,
        "creator" => :creator,
        "created" => :created,
        "modified" => :modified
      }.freeze

      def initialize
        @properties = {}
        @current_field = nil
        @text_buffer = +""
      end

      def start_element(_uri, local_name, qname, _attributes)
        name = element_name(local_name, qname)
        return unless FIELD_MAP.key?(name)

        @current_field = FIELD_MAP[name]
        @text_buffer = +""
      end

      def characters(text)
        @text_buffer << text if @current_field
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        return unless @current_field && FIELD_MAP.key?(name)

        @properties[@current_field] = @text_buffer.dup unless @text_buffer.empty?
        @current_field = nil
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener that captures cell style index (s attribute) from worksheet.
    class CellStyleListener
      include REXML::SAX2Listener

      attr_reader :cell_style_indices

      def initialize
        @cell_style_indices = {}
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        return unless name == "c"

        ref = attributes["r"]
        s = attributes["s"]
        @cell_style_indices[ref] = s.to_i if ref && s
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <sheetPr> element (tabColor, outlinePr).
    class SheetPropertiesListener
      include REXML::SAX2Listener

      attr_reader :properties

      def initialize
        @properties = {}
        @inside_sheet_pr = false
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        case name
        when "sheetPr"
          @inside_sheet_pr = true
        when "tabColor"
          @properties[:tab_color] = attributes["rgb"] if @inside_sheet_pr && attributes["rgb"]
        when "outlinePr"
          if @inside_sheet_pr
            sb = attributes["summaryBelow"]
            @properties[:summary_below] = %w[1 true].include?(sb) unless sb.nil?
            sr = attributes["summaryRight"]
            @properties[:summary_right] = %w[1 true].include?(sr) unless sr.nil?
          end
        end
      end

      def end_element(_uri, local_name, qname)
        name = element_name(local_name, qname)
        @inside_sheet_pr = false if name == "sheetPr"
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end

    # SAX2 listener for parsing <dimension> element.
    class DimensionListener
      include REXML::SAX2Listener

      attr_reader :ref

      def initialize
        @ref = nil
      end

      def start_element(_uri, local_name, qname, attributes)
        name = element_name(local_name, qname)
        @ref = attributes["ref"] if name == "dimension"
      end

      private

      def element_name(local_name, qname)
        if local_name.nil? || local_name.empty?
          qname.to_s.split(":").last
        else
          local_name
        end
      end
    end
  end
end
