# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Represents a single fully parsed, in-memory worksheet in a workbook.
    # Provides coordinate random access (sheet["A1"]), row lookups (row_at),
    # and immutable cell updates (update_cell).
    #
    # @example Access cells and rows
    #   sheet = workbook.sheet(0).load
    #   cell = sheet["A1"]
    #   row = sheet.row_at(0)
    #
    # @api public
    class Worksheet
      [Enumerable, CoordinateAccess].each { |m| include m }

      attr_reader :name, :rows, :columns, :charts, :conditional_formatting, :data_validations, :unmapped_data, :errors, :state, :hyperlinks, :comments, :styles, :date1904
      attr_writer :dimension
      alias conditional_formats conditional_formatting

      # @param name [String] The worksheet name (max 31 characters).
      # @param rows [Array<Elements::Row>] Rows in the sheet.
      # @param columns [Array<Elements::Column>] Column definitions.
      # @param charts [Array<Hash>] Charts in the sheet.
      # @param conditional_formatting [Array<Hash>, nil] Conditional formatting rules.
      # @param data_validations [Array<Hash>] Data validation rules.
      # @param unmapped_data [Hash] Additional metadata for round-tripping.
      # @param errors [Array<String>, nil] Validation errors.
      # @param conditional_formats [Array<Hash>, nil] Alias for conditional_formatting.
      # @param state [Symbol] Sheet visibility state (:visible, :hidden, or :very_hidden).
      # @param hyperlinks [Hash{String => Hash}] Hyperlink mappings by cell reference.
      # @param comments [Array<Hash>] Comments and notes in the sheet.
      # @param styles [Hash, nil] Optional parsed styles hash.
      # @param dimension [String, nil] Sheet dimension reference string (e.g. "A1:Z50000").
      # @param trim_empty_rows [Boolean] Whether to omit trailing empty rows during iteration.
      #: (name: String?, ?rows: Array[Elements::Row], ?columns: Array[Elements::Column], ?charts: Array[Hash[Symbol, untyped]], ?conditional_formatting: Array[Hash[Symbol, untyped]]?, ?data_validations: Array[Hash[Symbol, untyped]], ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?, ?conditional_formats: Array[Hash[Symbol, untyped]]?, ?state: Symbol, ?hyperlinks: Hash[String, Hash[Symbol, untyped]], ?comments: Array[Hash[Symbol, untyped]], ?styles: Hash[untyped, untyped]?, ?date1904: bool, ?dimension: String?, ?trim_empty_rows: bool) -> void
      def initialize(name:, rows: [], columns: [], charts: [], conditional_formatting: nil, data_validations: [], unmapped_data: {}, errors: nil, conditional_formats: nil, state: :visible, hyperlinks: {}, comments: [], styles: nil, date1904: false, dimension: nil, trim_empty_rows: false)
        @name = name
        @rows = (rows || []).freeze
        @columns = (columns || []).freeze
        @charts = (charts || []).freeze
        cf = conditional_formatting || conditional_formats || []
        @conditional_formatting = cf.freeze
        @data_validations = (data_validations || []).freeze
        @unmapped_data = (unmapped_data || {}).freeze
        computed_errors = errors || self.class.validate(@name, @rows)
        @errors = computed_errors.freeze
        @state = state ? state.to_sym : :visible
        @hyperlinks = (hyperlinks || {}).freeze
        @comments = (comments || []).freeze
        @styles = styles
        @date1904 = date1904 ? true : false
        @dimension = dimension
        @trim_empty_rows = trim_empty_rows ? true : false
      end

      # Returns whether trailing empty rows are omitted during iteration.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def trim_empty_rows?
        @trim_empty_rows ? true : false
      end
      alias trim_empty_rows trim_empty_rows?

      # Returns whether the worksheet uses the 1904 date system.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date1904?
        @date1904
      end

      # Returns the worksheet dimension reference string (e.g. "A1:Z50000"),
      # or computes it dynamically from rows if not explicitly set.
      #
      # @return [String, nil]
      # @api public
      #: () -> String?
      def dimension
        @dimension ||= compute_dimension
      end

      # Iterate over rows in the worksheet.
      #
      # @example
      #   sheet.each do |row|
      #     puts row.to_a.inspect
      #   end
      #
      # @yield [row]
      # @yieldparam row [Elements::Row]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::Row) -> void } -> void
      #: () -> Enumerator[Elements::Row, void]
      def each(&)
        return to_enum(:each) unless block_given?

        rows.each(&)
      end

      # Iterate over rows in the worksheet.
      #
      # @example
      #   sheet.each_row do |row|
      #     puts "Row #{row.index}: #{row.to_a.inspect}"
      #   end
      #
      # @overload each_row(trim_empty_rows: nil, &block)
      #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows.
      #   @yield [row]
      #   @yieldparam row [Elements::Row]
      #   @return [void]
      #
      # @overload each_row(trim_empty_rows: nil)
      #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows.
      #   @return [Enumerator, void]
      #
      # @api public
      #: (?trim_empty_rows: bool?) { (Elements::Row) -> void } -> void
      #: (?trim_empty_rows: bool?) -> Enumerator[Elements::Row, void]
      def each_row(trim_empty_rows: nil, &block)
        return to_enum(:each_row, trim_empty_rows: trim_empty_rows) unless block

        trim = if trim_empty_rows.nil?
                 @trim_empty_rows
               else
                 (trim_empty_rows ? true : false)
               end

        if trim
          pending_empty = []
          rows.each do |row|
            if row.empty?
              pending_empty << row
            else
              pending_empty.each(&block)
              pending_empty.clear
              block.call(row)
            end
          end
        else
          rows.each(&block)
        end
      end

      # Iterates over row values as Arrays.
      #
      # @overload each_row_values(type_cast: false, trim_empty_rows: nil, &block)
      #   @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
      #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows.
      #   @yield [values]
      #   @yieldparam values [Array<Object>] Row values array.
      #   @return [void]
      #
      # @overload each_row_values(type_cast: false, trim_empty_rows: nil)
      #   @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
      #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows.
      #   @return [Enumerator<Array<Object>, void>]
      #
      # @api public
      #: (?type_cast: bool, ?trim_empty_rows: bool?) { (Array[untyped]) -> void } -> void
      #: (?type_cast: bool, ?trim_empty_rows: bool?) -> Enumerator[Array[untyped], void]
      def each_row_values(type_cast: false, trim_empty_rows: nil, &block)
        return to_enum(:each_row_values, type_cast: type_cast, trim_empty_rows: trim_empty_rows) unless block

        trim = if trim_empty_rows.nil?
                 @trim_empty_rows
               else
                 (trim_empty_rows ? true : false)
               end

        each_row(trim_empty_rows: trim) do |row|
          block.call(row.values(type_cast: type_cast))
        end
      end

      # Iterate over all cells across rows.
      #
      # @yield [cell]
      # @yieldparam cell [Elements::Cell]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::Cell) -> void } -> void
      #: () -> Enumerator[Elements::Cell, void]
      def each_cell(&)
        return to_enum(:each_cell) unless block_given?

        cells.each(&)
      end

      # Returns whether the worksheet is valid according to OOXML specifications.
      #
      # @return [Boolean]
      #: () -> bool
      def valid?
        errors.empty?
      end

      # Returns whether the worksheet is hidden (:hidden or :very_hidden).
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hidden?
        @state == :hidden || @state == :very_hidden
      end

      # Returns whether the worksheet is visible.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def visible?
        @state == :visible
      end

      # Returns hyperlink metadata for the given cell reference, or nil.
      #
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
      # @param col [Integer, nil] Optional 0-based column index if ref_or_row is a row index.
      # @return [Hash, nil]
      # @api public
      #: (String | Symbol | Integer ref_or_row, ?Integer? col) -> Hash[Symbol, untyped]?
      def hyperlink(ref_or_row, col = nil)
        ref = if col
                "#{Cell.column_letter(col)}#{ref_or_row.to_i + 1}"
              else
                ref_or_row.to_s.upcase
              end
        @hyperlinks[ref]
      end

      # Returns comment metadata for the given cell reference, or nil.
      #
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
      # @param col [Integer, nil] Optional 0-based column index if ref_or_row is a row index.
      # @return [Hash, nil]
      # @api public
      #: (String | Symbol | Integer ref_or_row, ?Integer? col) -> Hash[Symbol, untyped]?
      def comment(ref_or_row, col = nil)
        ref = if col
                "#{Cell.column_letter(col)}#{ref_or_row.to_i + 1}"
              else
                ref_or_row.to_s.upcase
              end
        comments_by_ref[ref]
      end

      # Returns comments indexed by cell reference.
      #
      # @return [Hash<String, Hash[Symbol, untyped]>]
      # @api public
      #: () -> Hash[String, Hash[Symbol, untyped]]
      def comments_by_ref
        @comments_by_ref ||= @comments.each_with_object({}) do |c, acc|
          r = c[:ref] || c[:cell]
          acc[r.to_s.upcase] = c if r
        end
      end

      # Access a cell by reference or 0-based coordinates.
      #
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
      # @param col [Integer, nil] Optional 0-based column index.
      # @return [Elements::Cell, nil]
      # @api public
      #: (String | Symbol | Integer ref_or_row, ?Integer? col) -> Elements::Cell?
      def cell(ref_or_row, col = nil)
        if col
          row = row_at(ref_or_row.to_i)
          row&.cell_at(col.to_i)
        else
          self[ref_or_row]
        end
      end

      # Returns the formatted string representation of a cell's value.
      #
      # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
      # @param col [Integer, nil] Optional 0-based column index.
      # @return [String, nil]
      # @api public
      #: (String | Symbol | Integer ref_or_row, ?Integer? col) -> String?
      def formatted_value(ref_or_row, col = nil)
        cell(ref_or_row, col)&.formatted_value
      end

      # Returns a new Worksheet with the specified cell updated.
      #
      # @example
      #   new_sheet = sheet.update_cell("B1", value: "Updated")
      #
      # @param ref [String] The cell reference (e.g. "B1").
      # @param value [Object] The new cell value.
      # @param style_index [Integer, String, nil] Optional new style index.
      # @param formula [Elements::Formula, nil] Optional new formula.
      # @param hyperlink [Hash, String, nil] Optional new hyperlink.
      # @param comment [Hash, String, nil] Optional new comment.
      # @return [Worksheet] A new Worksheet instance.
      # @api public
      #: (String ref, ?value: untyped, ?style_index: Integer | String | nil, ?formula: Elements::Formula?, ?hyperlink: (Hash[Symbol, untyped] | String)?, ?comment: (Hash[Symbol, untyped] | String)?) -> Elements::Worksheet
      def update_cell(ref, value: nil, style_index: nil, formula: nil, hyperlink: nil, comment: nil)
        parsed = Cell.parse_ref(ref)
        raise ArgumentError, "invalid cell reference: #{ref}" unless parsed

        row_idx, col_idx = parsed
        existing_row = row_at(row_idx)

        if existing_row
          existing_cell = existing_row.cell_at(col_idx)
          new_cell = if existing_cell
                       existing_cell.with(
                         value: value || existing_cell.value,
                         style_index: style_index || existing_cell.style_index,
                         formula: formula || existing_cell.formula,
                         hyperlink: hyperlink || existing_cell.hyperlink,
                         comment: comment || existing_cell.comment
                       )
                     else
                       Cell.new(row_index: row_idx, column_index: col_idx, value: value, style_index: style_index, formula: formula, hyperlink: hyperlink, comment: comment)
                     end

          # Replace cell in the existing row
          new_cells = existing_row.cells.reject { |c| c.column_index == col_idx }
          new_cells << new_cell
          new_cells.sort_by!(&:column_index)

          new_row = existing_row.with(cells: new_cells)
          new_rows = rows.map { |r| r.index == row_idx ? new_row : r }
        else
          # Row doesn't exist, create it
          new_cell = Cell.new(row_index: row_idx, column_index: col_idx, value: value, style_index: style_index, formula: formula, hyperlink: hyperlink, comment: comment)
          new_row = Row.new(index: row_idx, cells: [new_cell])
          new_rows = (rows + [new_row]).sort_by!(&:index)
        end
        new_hyperlinks = hyperlink ? hyperlinks.merge(ref.upcase => hyperlink) : hyperlinks
        new_comments = if comment
                         entry = comment.is_a?(Hash) ? comment.merge(ref: ref.upcase) : { ref: ref.upcase, text: comment }
                         @comments.reject { |c| (c[:ref] || c[:cell])&.to_s&.upcase == ref.upcase } + [entry]
                       else
                         @comments
                       end
        with(rows: new_rows, hyperlinks: new_hyperlinks, comments: new_comments)
      end

      # Returns a new Worksheet with attributes replaced (Data-like behavior).
      #
      # @param changes [Hash]
      # @return [Worksheet]
      # @api public
      #: (**untyped) -> Elements::Worksheet
      def with(**changes)
        new_name = changes.key?(:name) ? changes[:name] : name
        new_rows = changes.key?(:rows) ? changes[:rows] : rows
        new_cols = changes.key?(:columns) ? changes[:columns] : columns
        new_charts = changes.key?(:charts) ? changes[:charts] : charts
        new_cf = if changes.key?(:conditional_formatting)
                   changes[:conditional_formatting]
                 elsif changes.key?(:conditional_formats)
                   changes[:conditional_formats]
                 else
                   conditional_formatting
                 end
        new_dv = changes.key?(:data_validations) ? changes[:data_validations] : data_validations
        new_unmapped = changes.key?(:unmapped_data) ? changes[:unmapped_data] : unmapped_data
        new_errors = changes.key?(:errors) ? changes[:errors] : errors
        new_state = changes.key?(:state) ? changes[:state] : state
        new_hyperlinks = changes.key?(:hyperlinks) ? changes[:hyperlinks] : hyperlinks
        new_comments = changes.key?(:comments) ? changes[:comments] : comments
        new_styles = changes.key?(:styles) ? changes[:styles] : styles
        new_date1904 = changes.key?(:date1904) ? changes[:date1904] : date1904
        new_dim = changes.key?(:dimension) ? changes[:dimension] : @dimension
        new_trim = changes.key?(:trim_empty_rows) ? changes[:trim_empty_rows] : @trim_empty_rows

        self.class.new(
          name: new_name,
          rows: new_rows,
          columns: new_cols,
          charts: new_charts,
          conditional_formatting: new_cf,
          data_validations: new_dv,
          unmapped_data: new_unmapped,
          errors: new_errors,
          state: new_state,
          hyperlinks: new_hyperlinks,
          comments: new_comments,
          styles: new_styles,
          date1904: new_date1904,
          dimension: new_dim,
          trim_empty_rows: new_trim
        )
      end

      # Support pattern matching.
      #: (Array[Symbol]?) -> Hash[Symbol, untyped]
      def deconstruct_keys(_keys)
        {
          name: name,
          rows: rows,
          columns: columns,
          charts: charts,
          conditional_formatting: conditional_formatting,
          conditional_formats: conditional_formatting,
          data_validations: data_validations,
          unmapped_data: unmapped_data,
          errors: errors,
          state: state,
          hyperlinks: hyperlinks,
          comments: comments,
          styles: styles,
          date1904: date1904,
          dimension: dimension,
          trim_empty_rows: @trim_empty_rows
        }
      end

      # Compare worksheets for equality.
      #: (untyped other) -> bool
      def ==(other)
        return false unless other.is_a?(Worksheet)

        name == other.name && rows == other.rows && columns == other.columns && charts == other.charts &&
          conditional_formatting == other.conditional_formatting && data_validations == other.data_validations &&
          state == other.state && hyperlinks == other.hyperlinks && comments == other.comments
      end
      alias eql? ==

      #: () -> Integer
      def hash
        [self.class, name, rows, columns, charts, conditional_formatting, data_validations, state, hyperlinks, comments].hash
      end

      # Returns self when load is called on an already in-memory Worksheet.
      #
      # @return [Elements::Worksheet]
      # @api public
      #: () -> Elements::Worksheet
      def load
        self
      end
      alias to_worksheet load

      FORBIDDEN_NAME_CHARS = %r{[\\/?*\[\]]}

      # Returns whether the worksheet name is valid according to OOXML specifications (1..31 chars, no forbidden chars).
      #
      # @param name [Object]
      # @return [Boolean]
      #: (untyped name) -> bool
      def self.valid_name?(name)
        case name
        when String
          !name.empty? && name.size <= 31 && !name.match?(FORBIDDEN_NAME_CHARS)
        else
          false
        end
      end

      # Validates worksheet name and rows against OOXML limits.
      #
      # @param name [String]
      # @param rows [Array<Elements::Row>]
      # @return [Array<String>] List of errors.
      #: (untyped name, untyped rows) -> Array[String]
      def self.validate(name, rows)
        errs = []
        if name.nil? || !name.is_a?(String) || name.empty?
          errs << "worksheet name must be a non-empty String (got #{name.inspect})"
        else
          errs << "worksheet name cannot exceed 31 characters (got #{name.size})" if name.size > 31
          errs << "worksheet name cannot contain \\, /, ?, *, [, or ]" if name.match?(FORBIDDEN_NAME_CHARS)
        end
        errs << "rows must be an Array (got #{rows.class})" unless rows.is_a?(Array)
        if rows.is_a?(Array)
          indices = rows.map(&:index)
          if indices.uniq.size != indices.size
            dups = indices.select { |i| indices.count(i) > 1 }.uniq
            errs << "duplicate row index: #{dups.join(", ")} — row indices within a sheet must be unique"
          end
        end
        errs
      end

      private

      #: () -> String?
      def compute_dimension
        return nil if rows.empty?

        min_r = nil
        max_r = nil
        min_c = nil
        max_c = nil

        rows.each do |r|
          next if r.cells.empty?

          min_r = r.index if min_r.nil? || r.index < min_r
          max_r = r.index if max_r.nil? || r.index > max_r
          r.cells.each do |c|
            min_c = c.column_index if min_c.nil? || c.column_index < min_c
            max_c = c.column_index if max_c.nil? || c.column_index > max_c
          end
        end

        return nil if min_r.nil? || min_c.nil?

        start_cell = "#{Elements::Cell.column_letter(min_c)}#{min_r + 1}"
        end_cell = "#{Elements::Cell.column_letter(max_c)}#{max_r + 1}"
        start_cell == end_cell ? start_cell : "#{start_cell}:#{end_cell}"
      end
    end
  end
end
