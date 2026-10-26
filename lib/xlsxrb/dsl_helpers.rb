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

    # Normalizes table options hash.
    #
    # @param ref [String]
    # @param columns [Array<String>, Array<Hash>]
    # @param name [String, nil]
    # @param display_name [String, nil]
    # @param style [String, nil]
    # @param opts [Hash]
    # @return [Hash]
    #: (String ref, columns: untyped, ?name: String?, ?display_name: String?, ?style: String?, **untyped opts) -> Hash[Symbol, untyped]
    def self.normalize_table_options(ref, columns:, name: nil, display_name: nil, style: nil, **opts)
      tbl = { ref: ref, columns: columns }
      tbl[:name] = name if name
      tbl[:display_name] = display_name if display_name
      tbl[:style] = style if style
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

    # Compresses an array of integers or ranges into a minimal sorted list of disjoint [first, last] inclusive pairs.
    #
    # @param values [Integer, Range, Enumerable, nil]
    # @return [Array<Array(Integer, Integer)>]
    #: (untyped values) -> Array[[Integer, Integer]]
    def self.compress_runs(values)
      return [] if values.nil?

      intervals = []
      if values.is_a?(Integer)
        intervals << [values, values]
      elsif values.is_a?(Range)
        first = values.begin.to_i
        last = values.exclude_end? ? values.end.to_i - 1 : values.end.to_i
        intervals << [first, last] if first <= last
      else
        values.each do |item|
          if item.is_a?(Range)
            first = item.begin.to_i
            last = item.exclude_end? ? item.end.to_i - 1 : item.end.to_i
            intervals << [first, last] if first <= last
          elsif item.is_a?(Integer)
            intervals << [item, item]
          end
        end
      end

      return [] if intervals.empty?

      intervals.sort_by!(&:first)
      merged = []
      intervals.each do |s, e|
        if !merged.empty? && s <= merged.last[1] + 1
          merged.last[1] = [merged.last[1], e].max
        else
          merged << [s, e]
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
        next if e_end < start_pt
        break if e_start > max

        runs << [start_pt, e_start - 1] if e_start > start_pt
        start_pt = [start_pt, e_end + 1].max
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
      c1 = Elements::Cell.column_letter(col_run[0])
      r1 = row_run[0]
      c2 = Elements::Cell.column_letter(col_run[1])
      r2 = row_run[1]
      if col_run[0] == col_run[1] && r1 == r2
        "#{c1}#{r1}"
      else
        "#{c1}#{r1}:#{c2}#{r2}"
      end
    end

    # Combines column intervals and row intervals into chunked space-delimited sqref strings
    # respecting the maximum reference limit per dataValidation tag (default 32,767).
    #
    # @param col_runs [Array<Array(Integer, Integer)>]
    # @param row_runs [Array<Array(Integer, Integer)>]
    # @param max_per_validation [Integer]
    # @return [Array<String>]
    #: (Array[Array[Integer]] col_runs, Array[Array[Integer]] row_runs, ?Integer max_per_validation) -> Array[String]
    def self.chunked_sqrefs(col_runs, row_runs, max_per_validation = Xlsxrb::MAX_REFS_PER_VALIDATION)
      areas = col_runs.flat_map do |col_run|
        row_runs.map { |row_run| area_ref(col_run, row_run) }
      end
      return [] if areas.empty?

      areas.each_slice(max_per_validation).map { |chunk| chunk.join(" ") }
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
  end
end
