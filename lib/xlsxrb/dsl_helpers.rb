# frozen_string_literal: true

# rbs_inline: enabled

require_relative "elements/cell"
require_relative "ooxml/utils"

module Xlsxrb
  # Internal helper module for normalizing and validating DSL arguments
  # shared between {WorksheetBuilder} (in-memory) and {StreamWriter} (streaming).
  #
  # @api private
  module DslHelpers
    # Normalizes merge cell range string or coordinate components into a canonical range reference (e.g. "A1:C3").
    #
    # @param range [String, Hash, nil]
    # @param row [Integer, nil]
    # @param col_start [Integer, String, nil]
    # @param col_end [Integer, String, nil]
    # @param row_start [Integer, nil]
    # @param row_end [Integer, nil]
    # @param strict_excel_mode [Boolean]
    # @return [String]
    #: (untyped range, ?row: Integer?, ?col_start: (Integer | String)?, ?col_end: (Integer | String)?, ?row_start: Integer?, ?row_end: Integer?, ?strict_excel_mode: bool) -> String
    def self.normalize_merge_range(range = nil, row: nil, col_start: nil, col_end: nil, row_start: nil, row_end: nil, strict_excel_mode: true)
      if range.is_a?(Hash)
        row = range[:row]
        row_start = range[:row_start]
        row_end = range[:row_end]
        col_start = range[:col_start]
        col_end = range[:col_end]
        range = nil
      end

      if range
        raise ArgumentError, "Invalid merge range format: '#{range}'. Expected format like 'A1:B2'." if strict_excel_mode && !range.match?(/\A[A-Za-z]{1,3}\d+(:[A-Za-z]{1,3}\d+)?\z/)

        range
      else
        r_start = row || row_start || 0
        r_end = row || row_end || 0
        c_start = Elements::Cell.column_index(col_start || 0)
        c_end = Elements::Cell.column_index(col_end || 0)
        start_ref = "#{Elements::Cell.column_letter(c_start)}#{r_start + 1}"
        end_ref = "#{Elements::Cell.column_letter(c_end)}#{r_end + 1}"
        "#{start_ref}:#{end_ref}"
      end
    end

    # Normalizes sheet protection options and hashes plain-text password if supplied.
    #
    # @param opts [Hash]
    # @return [Hash]
    #: (Hash[Symbol, untyped] opts) -> Hash[Symbol, untyped]
    def self.normalize_protection_options(opts)
      normalized = opts.dup
      plain_password = normalized[:password]
      needs_hash = plain_password.is_a?(String) && !plain_password.empty? &&
                   normalized[:algorithm_name].nil? && normalized[:hash_value].nil? &&
                   normalized[:salt_value].nil? && normalized[:spin_count].nil? &&
                   !plain_password.match?(/\A[0-9A-Fa-f]{4}\z/)
      if needs_hash
        normalized.delete(:password)
        normalized.merge!(Ooxml::Utils.hash_password(plain_password))
      end
      normalized
    end

    # Normalizes column index or letter or Range to an array of 0-based column integers.
    #
    # @param index [Integer, String, Symbol, Range, Array]
    # @return [Array<Integer>]
    #: (untyped index) -> Array[Integer]
    def self.normalize_column_indices(index)
      case index
      when Range, Array
        index.map { |i| Elements::Cell.column_index(i) }
      else
        [Elements::Cell.column_index(index)]
      end
    end

    # Normalizes page margins into a compact hash.
    #
    # @param left [Float, nil]
    # @param right [Float, nil]
    # @param top [Float, nil]
    # @param bottom [Float, nil]
    # @param header [Float, nil]
    # @param footer [Float, nil]
    # @return [Hash{Symbol => Float}]
    #: (?left: Float?, ?right: Float?, ?top: Float?, ?bottom: Float?, ?header: Float?, ?footer: Float?) -> Hash[Symbol, Float]
    def self.normalize_page_margins(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
      { left: left, right: right, top: top, bottom: bottom, header: header, footer: footer }.compact
    end

    TOTALS_ROW_FUNCTIONS = {
      sum: "sum",
      average: "average",
      count: "count",
      count_nums: "countNums",
      countnums: "countNums",
      max: "max",
      min: "min",
      std_dev: "stdDev",
      stddev: "stdDev",
      var: "var",
      custom: "custom"
    }.freeze

    # Normalizes table options hash.
    #
    # @param ref [String]
    # @param columns [Array<String>, Array<Hash>, Array<Symbol>]
    # @param name [String, nil]
    # @param display_name [String, nil]
    # @param style [String, Hash, nil]
    # @param opts [Hash]
    # @return [Hash]
    #: (String ref, columns: untyped, ?name: String?, ?display_name: String?, ?style: (String | Hash[Symbol, untyped])?, **untyped opts) -> Hash[Symbol, untyped]
    def self.normalize_table_options(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      tbl = { ref: ref }
      cols = columns.is_a?(Array) ? columns : [columns]
      tbl[:columns] = cols.map do |c|
        if c.is_a?(Hash)
          col = c.dup
          col[:name] = col[:name].to_s if col.key?(:name)
          if col.key?(:totals_row_function)
            fn = col[:totals_row_function]
            col[:totals_row_function] = TOTALS_ROW_FUNCTIONS[fn.to_sym] || fn.to_s
          end
          col
        else
          c.to_s
        end
      end
      tbl[:name] = name.to_s if name
      tbl[:display_name] = display_name.to_s if display_name
      tbl[:style] = style if style

      # Normalize total row options
      if opts.key?(:total_row) || opts.key?(:totals_row)
        tr = opts.delete(:total_row)
        tr = opts.delete(:totals_row) if tr.nil?
        tbl[:totals_row_count] = if tr == true
                                   1
                                 elsif tr == false
                                   0
                                 else
                                   tr.to_i
                                 end
      elsif opts.key?(:total_row_count)
        tbl[:totals_row_count] = opts.delete(:total_row_count).to_i
      elsif opts.key?(:totals_row_count)
        tbl[:totals_row_count] = opts.delete(:totals_row_count).to_i
      end

      tbl.merge!(opts)
      tbl
    end

    # Converts a cell range (e.g. "A1:B2" or "A1") to an absolute reference format (e.g. "$A$1:$B$2" or "$A$1").
    #
    # @param range [String]
    # @return [String]
    #: (String range) -> String
    def self.absolute_range(range)
      range.gsub(/(?:\$)?([A-Z]+)(?:\$)?(\d+)\b/, '$\1$\2')
    end

    # Normalizes row values: if given a column-keyed Hash, converts to a sparse Array.
    # If already an Array or other non-Hash enumerable, returns as-is without allocations.
    #
    # @param values [Array<Object>, Hash{String, Integer, Symbol => Object}, nil]
    # @return [Array<Object>, nil]
    #: (untyped values) -> untyped
    def self.normalize_row_values(values)
      return values unless values.is_a?(Hash)

      cells_array = []
      values.each do |k, v|
        cells_array[Elements::Cell.column_index(k)] = v
      end
      cells_array
    end

    # Normalizes row styles: if given a Hash of column-keyed styles (including Ranges/Arrays),
    # expands to a sparse Array of styles. If already an Array or nil, returns as-is without allocations.
    #
    # @param styles [untyped]
    # @return [untyped]
    #: (untyped styles) -> untyped
    def self.normalize_row_styles(styles)
      return styles unless styles.is_a?(Hash)

      styles_array = []
      styles.each do |k, v|
        if k.is_a?(Range) || k.is_a?(Array)
          k.each { |idx| styles_array[Elements::Cell.column_index(idx)] = v }
        else
          styles_array[Elements::Cell.column_index(k)] = v
        end
      end
      styles_array
    end

    # Validates row index and height against Excel limits in strict mode.
    #
    # @param row_index [Integer]
    # @param height [Float, Integer, nil]
    # @param strict_excel_mode [Boolean]
    # @return [void]
    #: (Integer row_index, Float | Integer | nil height, ?strict_excel_mode: bool) -> void
    def self.validate_row_bounds!(row_index, height, strict_excel_mode: true)
      return unless strict_excel_mode

      raise ArgumentError, "Row index #{row_index} exceeds Excel limit of 1,048,576 rows" if row_index >= 1_048_576
      raise ArgumentError, "Row height #{height} must be between 0 and 409 points (Excel limitation)" if height && (height.negative? || height > 409)
    end

    # Compresses a sorted or unsorted collection of integers and/or ranges into minimal disjoint [start, finish] runs.
    # Operates on intervals directly without expanding ranges into individual integer arrays.
    #
    # @param values [Integer, Range, Enumerable, nil]
    # @return [Array<Array(Integer, Integer)>]
    #: (untyped values) -> Array[[Integer, Integer]]
    def self.compress_runs(values)
      intervals = []
      items = values.is_a?(Array) ? values : [values]
      items.each do |item|
        case item
        when Range
          first = item.begin
          last = item.exclude_end? ? item.end - 1 : item.end
          intervals << [first, last] if first <= last
        when Integer
          intervals << [item, item]
        end
      end

      merged = []
      intervals.sort.each do |s, e|
        last = merged.last
        if !last || s > last.last + 1
          merged << [s, e]
        else
          last[1] = [last.last, e].max
        end
      end
      merged
    end

    # Returns minimal disjoint runs within min..max that are NOT present in excluded.
    #
    # @param excluded [Integer, Range, Enumerable, nil]
    # @param min [Integer]
    # @param max [Integer]
    # @return [Array<Array(Integer, Integer)>]
    #: (untyped excluded, Integer min, Integer max) -> Array[[Integer, Integer]]
    def self.complement_runs(excluded, min, max)
      runs = []
      start_pt = min
      compress_runs(excluded).each do |e_start, e_end|
        if e_end >= start_pt && e_start <= max
          runs << [start_pt, e_start - 1] if e_start > start_pt
          start_pt = e_end + 1
        end
      end
      runs << [start_pt, max] if start_pt <= max
      runs
    end

    # Builds an A1-style reference from a column interval [col_start, col_end] (0-based)
    # and a row interval [row_start, row_end] (1-based Excel row number).
    #
    # @param col_run [Array(Integer, Integer)]
    # @param row_run [Array(Integer, Integer)]
    # @return [String]
    #: (Array[Integer] col_run, Array[Integer] row_run) -> String
    def self.area_ref(col_run, row_run)
      c1, c2 = col_run
      r1, r2 = row_run
      ref1 = "#{Elements::Cell.column_letter(c1)}#{r1}"
      c1 == c2 && r1 == r2 ? ref1 : "#{ref1}:#{Elements::Cell.column_letter(c2)}#{r2}"
    end

    # Combines column intervals and row intervals into chunked space-delimited sqref strings
    # respecting the maximum reference limit per dataValidation tag (default 32,767).
    #
    # @param col_runs [Array<Array(Integer, Integer)>]
    # @param row_runs [Array<Array(Integer, Integer)>]
    # @param max_per_validation [Integer]
    # @return [Array<String>]
    #: (Array[Array[Integer]] col_runs, Array[Array[Integer]] row_runs, ?Integer max_per_validation) -> Array[String]
    def self.chunked_sqrefs(col_runs, row_runs, max_per_validation = MAX_REFS_PER_VALIDATION)
      refs = col_runs.flat_map { |col_run| row_runs.map { |row_run| area_ref(col_run, row_run) } }
      refs.each_slice(max_per_validation).map { |chunk| chunk.join(" ") }
    end

    # Quotes a sheet name for use in formula and defined name references, escaping single quotes with double single quotes.
    #
    # @param sheet_name [String, Object]
    # @return [String]
    #: (untyped sheet_name) -> String
    def self.quote_sheet_name(sheet_name)
      "'#{sheet_name.to_s.gsub("'", "''")}'"
    end

    # Converts a cell reference, rectangular cell range, column range, or row range to absolute $ form.
    # Returns nil if the input is not a coordinate or range.
    #
    # @param ref [String]
    # @return [String, nil]
    #: (String ref) -> String?
    def self.format_absolute_reference(ref)
      u = ref.upcase
      case u
      when /\A\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?\z/
        absolute_range(u)
      when /\A\$?([A-Z]{1,3}):\$?([A-Z]{1,3})\z/, /\A\$?(\d+):\$?(\d+)\z/
        "$#{Regexp.last_match(1)}:$#{Regexp.last_match(2)}"
      end
    end

    # Normalizes defined name value, auto-formatting bare coordinates and ranges with sheet quotes and absolute dollar references.
    # Leaves constants, formulas, and already-qualified references untouched.
    #
    # @param value [Object, nil]
    # @param sheet [String, nil]
    # @param col [Integer, String, nil]
    # @param rows [Integer, Range, nil]
    # @param range [String, nil]
    # @return [String]
    #: (?untyped value, ?sheet: untyped, ?col: untyped, ?rows: untyped, ?range: untyped) -> String
    def self.normalize_defined_name_value(value = nil, sheet: nil, col: nil, rows: nil, range: nil)
      if range
        value = range
      elsif col || rows
        c_letter = Elements::Cell.column_letter(col ? Elements::Cell.column_index(col) : 0)
        value = if rows.is_a?(Range)
                  r2 = rows.exclude_end? ? rows.end - 1 : rows.end
                  "$#{c_letter}$#{rows.begin}:$#{c_letter}$#{r2}"
                elsif rows
                  "$#{c_letter}$#{rows}"
                else
                  "$#{c_letter}:$#{c_letter}"
                end
      end

      str = value.to_s
      abs = format_absolute_reference(str)
      if abs
        sheet ? "#{quote_sheet_name(sheet)}!#{abs}" : abs
      else
        str
      end
    end

    # Normalizes data validation arguments (sqref, coordinate targeting, list arrays, input/error message inference).
    #
    # @param sqref [String, nil]
    # @param opts [Hash]
    # @yield [DataValidationBuilder]
    # @return [Array<Hash>]
    #: (?String? sqref, **untyped opts) ?{ (DataValidationBuilder) -> void } -> Array[Hash[Symbol, untyped]]
    def self.normalize_data_validation(sqref = nil, **opts)
      if block_given?
        builder = DataValidationBuilder.new(sqref, **opts)
        yield builder
        opts = builder.to_h
      else
        opts = opts.dup
      end
      sqref = opts.delete(:sqref) || sqref

      if opts[:in].is_a?(Array)
        opts[:type] ||= "list"
        opts[:formula1] ||= %("#{opts.delete(:in).join(",")}")
      end

      # ST_Xstring conversion for newlines in prompt and error messages
      opts[:prompt] = opts[:prompt].to_s.gsub(/\r\n|\r|\n/, "_x000a_") if opts[:prompt]
      opts[:error] = opts[:error].to_s.gsub(/\r\n|\r|\n/, "_x000a_") if opts[:error]
      opts[:prompt_title] = opts[:prompt_title].to_s.gsub(/\r\n|\r|\n/, "_x000a_") if opts[:prompt_title]
      opts[:error_title] = opts[:error_title].to_s.gsub(/\r\n|\r|\n/, "_x000a_") if opts[:error_title]

      # Auto-infer show_input_message / show_error_message
      opts[:show_input_message] = true if !opts.key?(:show_input_message) && (!opts[:prompt].nil? || !opts[:prompt_title].nil?)
      opts[:show_error_message] = true if !opts.key?(:show_error_message) && (!opts[:error].nil? || !opts[:error_title].nil?)

      # Determine sqrefs
      sqrefs = []
      has_coords = opts.key?(:columns) || opts.key?(:col) || opts.key?(:rows) || opts.key?(:row)
      if sqref.nil? && has_coords
        cols_val = opts.delete(:columns)
        col_val = opts.delete(:col)
        raw_cols = cols_val || col_val || 0

        rows_val = opts.delete(:rows)
        row_val = opts.delete(:row)
        raw_rows = rows_val || row_val || (1..Xlsxrb::MAX_ROW)

        except_rows = opts.delete(:except_rows)

        col_indices = normalize_column_indices(raw_cols)
        col_runs = compress_runs(col_indices)
        row_runs = compress_runs(raw_rows)
        row_runs = row_runs.flat_map { |s, e| complement_runs(except_rows, s, e) } if except_rows
        sqrefs = chunked_sqrefs(col_runs, row_runs)
      elsif sqref
        opts.delete(:columns)
        opts.delete(:col)
        opts.delete(:rows)
        opts.delete(:row)
        opts.delete(:except_rows)
        str_sqref = sqref.to_s
        refs = str_sqref.split
        sqrefs = if refs.size > Xlsxrb::MAX_REFS_PER_VALIDATION
                   refs.each_slice(Xlsxrb::MAX_REFS_PER_VALIDATION).map { |chunk| chunk.join(" ") }
                 else
                   [str_sqref]
                 end
      end

      raise ArgumentError, "data validation sqref or coordinates (columns/rows) must be provided" if sqrefs.empty?

      sqrefs.map { |ref| opts.merge(sqref: ref) }
    end

    CF_OPERATORS = {
      greater_than: "greaterThan",
      less_than: "lessThan",
      equal: "equal",
      not_equal: "notEqual",
      greater_than_or_equal: "greaterThanOrEqual",
      less_than_or_equal: "lessThanOrEqual",
      between: "between",
      not_between: "notBetween",
      contains_text: "containsText",
      not_contains: "notContains",
      begins_with: "beginsWith",
      ends_with: "endsWith"
    }.freeze

    CF_TYPES = {
      cell_is: "cellIs",
      expression: "expression",
      color_scale: "colorScale",
      data_bar: "dataBar",
      icon_set: "iconSet",
      above_average: "aboveAverage",
      top10: "top10",
      duplicate_values: "duplicateValues",
      unique_values: "uniqueValues",
      contains_text: "containsText",
      not_contains_text: "notContainsText",
      begins_with: "beginsWith",
      ends_with: "endsWith",
      contains_blanks: "containsBlanks",
      not_contains_blanks: "notContainsBlanks",
      time_period: "timePeriod"
    }.freeze

    CF_ICON_SETS = {
      three_arrows: "3Arrows",
      three_arrows_gray: "3ArrowsGray",
      three_flags: "3Flags",
      three_traffic_lights1: "3TrafficLights1",
      three_traffic_lights2: "3TrafficLights2",
      three_signs: "3Signs",
      three_symbols: "3Symbols",
      three_symbols2: "3Symbols2",
      four_arrows: "4Arrows",
      four_arrows_gray: "4ArrowsGray",
      four_red_to_black: "4RedToBlack",
      four_rating: "4Rating",
      four_traffic_lights: "4TrafficLights",
      five_arrows: "5Arrows",
      five_arrows_gray: "5ArrowsGray",
      five_rating: "5Rating",
      five_quarters: "5Quarters"
    }.freeze

    # Extracts a differential format (dxf) hash from a conditional formatting rule hash.
    # Mutates the rule hash by deleting extracted font and fill keys.
    #
    # @param rule [Hash]
    # @return [Hash, nil]
    #: (Hash[Symbol, untyped] rule) -> Hash[Symbol, untyped]?
    def self.extract_dxf_from_rule!(rule)
      font = (rule.delete(:font) || {}).dup
      font[:color] = rule.delete(:font_color) if rule.key?(:font_color)
      font[:bold] = rule.delete(:bold) if rule.key?(:bold)
      font[:italic] = rule.delete(:italic) if rule.key?(:italic)
      font[:underline] = rule.delete(:underline) if rule.key?(:underline)
      font[:size] = rule.delete(:size) if rule.key?(:size)
      font[:sz] = rule.delete(:sz) if rule.key?(:sz)

      fill = (rule.delete(:fill) || {}).dup
      if rule.key?(:fill_color)
        fill[:pattern] = "solid"
        fill[:fg_color] = rule.delete(:fill_color)
      end

      dxf = {}
      dxf[:font] = font unless font.empty?
      dxf[:fill] = fill unless fill.empty?
      dxf.empty? ? nil : dxf
    end

    # Normalizes conditional format arguments (sqref, coordinate targeting, types, operators, color scales, icon sets).
    #
    # @param sqref [String, nil]
    # @param opts [Hash]
    # @yield [ConditionalFormatBuilder]
    # @return [Array<Hash>]
    #: (?String? sqref, **untyped opts) ?{ (ConditionalFormatBuilder) -> void } -> Array[Hash[Symbol, untyped]]
    def self.normalize_conditional_format(sqref = nil, **opts)
      if block_given?
        builder = ConditionalFormatBuilder.new(sqref, **opts)
        yield builder
        opts = builder.to_h
      else
        opts = opts.dup
      end
      sqref = opts.delete(:sqref) || sqref

      # Determine sqrefs
      sqrefs = []
      has_coords = opts.key?(:columns) || opts.key?(:col) || opts.key?(:rows) || opts.key?(:row)
      if sqref.nil? && has_coords
        cols_val = opts.delete(:columns)
        col_val = opts.delete(:col)
        raw_cols = cols_val || col_val || 0

        rows_val = opts.delete(:rows)
        row_val = opts.delete(:row)
        raw_rows = rows_val || row_val || (1..Xlsxrb::MAX_ROW)

        except_rows = opts.delete(:except_rows)

        col_indices = normalize_column_indices(raw_cols)
        col_runs = compress_runs(col_indices)
        row_runs = compress_runs(raw_rows)
        row_runs = row_runs.flat_map { |s, e| complement_runs(except_rows, s, e) } if except_rows
        sqrefs = chunked_sqrefs(col_runs, row_runs)
      elsif sqref
        opts.delete(:columns)
        opts.delete(:col)
        opts.delete(:rows)
        opts.delete(:row)
        opts.delete(:except_rows)
        str_sqref = sqref.to_s
        refs = str_sqref.split
        sqrefs = if refs.size > Xlsxrb::MAX_REFS_PER_VALIDATION
                   refs.each_slice(Xlsxrb::MAX_REFS_PER_VALIDATION).map { |chunk| chunk.join(" ") }
                 else
                   [str_sqref]
                 end
      end

      raise ArgumentError, "conditional format sqref or coordinates (columns/rows) must be provided" if sqrefs.empty?

      type_str = opts[:type].to_s
      # Normalize color_scale
      if opts.key?(:color_scale) || type_str == "colorScale" || type_str == "color_scale"
        cs = (opts[:color_scale] || {}).dup
        colors = cs[:colors] || opts.delete(:colors)
        if colors.is_a?(Array)
          colors = colors.map do |c|
            c.is_a?(String) && c =~ /\A[0-9A-Fa-f]{6}\z/ ? "FF#{c.upcase}" : c
          end
          cs[:colors] = colors
          if !cs.key?(:cfvo) || cs[:cfvo].nil?
            cs[:cfvo] = if colors.size == 2
                          [{ type: "min" }, { type: "max" }]
                        else
                          [{ type: "min" }, { type: "percent", val: "50" }, { type: "max" }]
                        end
          end
        end
        opts[:color_scale] = cs unless cs.empty?
      end

      # Normalize icon_set
      if opts.key?(:icon_set) || type_str == "iconSet" || type_str == "icon_set"
        is = (opts[:icon_set] || {}).dup
        style_name = is[:icon_set] || opts.delete(:icon_style) || is[:icon_style]
        if style_name
          mapped_style = CF_ICON_SETS[style_name.to_sym] || style_name.to_s
          is[:icon_set] = mapped_style
          if !is.key?(:cfvo) || is[:cfvo].nil?
            is[:cfvo] = if mapped_style.start_with?("4")
                          [
                            { type: "percent", val: "0" },
                            { type: "percent", val: "25" },
                            { type: "percent", val: "50" },
                            { type: "percent", val: "75" }
                          ]
                        elsif mapped_style.start_with?("5")
                          [
                            { type: "percent", val: "0" },
                            { type: "percent", val: "20" },
                            { type: "percent", val: "40" },
                            { type: "percent", val: "60" },
                            { type: "percent", val: "80" }
                          ]
                        else
                          [
                            { type: "percent", val: "0" },
                            { type: "percent", val: "33" },
                            { type: "percent", val: "67" }
                          ]
                        end
          end
        end
        opts[:icon_set] = is unless is.empty?
      end

      sqrefs.map { |ref| opts.merge(sqref: ref) }
    end

    FORBIDDEN_SHEET_CHARS = %r{[\\/?*\[\]:]|[\x00-\x08\x0B\x0C\x0E-\x1F]}
    MAX_SHEET_NAME_LENGTH = 31

    # Sanitizes a worksheet name to ensure it conforms to Excel constraints within 31 characters.
    # Replaces forbidden characters (\ / ? * [ ] : and control characters),
    # replaces leading/trailing apostrophes, and replaces reserved 'History' with 'History_'.
    #
    # @param name [Object] Desired worksheet name.
    # @param replace_char [String] Replacement character for invalid characters.
    # @return [String] Excel-compliant worksheet name.
    #: (untyped name, ?replace_char: String) -> String
    def self.sanitize_sheet_name(name, replace_char: "_")
      raw = name.to_s.strip
      str = raw.gsub(FORBIDDEN_SHEET_CHARS, replace_char)
      str = str.sub(/\A'+/) { replace_char * Regexp.last_match(0).length }
      str = "Sheet" if str.empty?

      base = str[0, MAX_SHEET_NAME_LENGTH].sub(/'+\z/) { replace_char * Regexp.last_match(0).length }
      return "#{base}#{replace_char}" if base.casecmp?("History")

      base
    end

    class << self
      alias sheet_name sanitize_sheet_name
    end
  end
end
