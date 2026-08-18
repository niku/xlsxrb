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

      attr_reader :name, :rows, :columns, :charts, :unmapped_data, :errors

      # @param name [String] The worksheet name (max 31 characters).
      # @param rows [Array<Elements::Row>] Rows in the sheet.
      # @param columns [Array<Elements::Column>] Column definitions.
      # @param charts [Array<Hash>] Charts in the sheet.
      # @param unmapped_data [Hash] Additional metadata for round-tripping.
      # @param errors [Array<String>, nil] Validation errors.
      #: (name: String, ?rows: Array[Elements::Row], ?columns: Array[Elements::Column], ?charts: Array[Hash[Symbol, untyped]], ?unmapped_data: Hash[untyped, untyped], ?errors: Array[String]?) -> void
      def initialize(name:, rows: [], columns: [], charts: [], unmapped_data: {}, errors: nil)
        @name = name
        @rows = (rows || []).freeze
        @columns = (columns || []).freeze
        @charts = (charts || []).freeze
        @unmapped_data = (unmapped_data || {}).freeze
        computed_errors = errors || self.class.validate(@name, @rows)
        @errors = computed_errors.freeze
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
      #: | () -> Enumerator[Elements::Row, void]
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
      # @yield [row]
      # @yieldparam row [Elements::Row]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::Row) -> void } -> void
      #: | () -> Enumerator[Elements::Row, void]
      def each_row(&)
        return to_enum(:each_row) unless block_given?

        rows.each(&)
      end

      # Iterate over all cells across rows.
      #
      # @yield [cell]
      # @yieldparam cell [Elements::Cell]
      # @return [Enumerator, void]
      # @api public
      #: () { (Elements::Cell) -> void } -> void
      #: | () -> Enumerator[Elements::Cell, void]
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

      # Returns a new Worksheet with the specified cell updated.
      #
      # @example
      #   new_sheet = sheet.update_cell("B1", value: "Updated")
      #
      # @param ref [String] The cell reference (e.g. "B1").
      # @param value [Object] The new cell value.
      # @param style_index [Integer, String, nil] Optional new style index.
      # @param formula [Elements::Formula, nil] Optional new formula.
      # @return [Worksheet] A new Worksheet instance.
      # @api public
      #: (String ref, ?value: untyped, ?style_index: Integer | String | nil, ?formula: Elements::Formula?) -> Elements::Worksheet
      def update_cell(ref, value: nil, style_index: nil, formula: nil)
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
                         formula: formula || existing_cell.formula
                       )
                     else
                       Cell.new(row_index: row_idx, column_index: col_idx, value: value, style_index: style_index, formula: formula)
                     end

          # Replace cell in the existing row
          new_cells = existing_row.cells.reject { |c| c.column_index == col_idx }
          new_cells << new_cell
          new_cells.sort_by!(&:column_index)

          new_row = existing_row.with(cells: new_cells)
          new_rows = rows.map { |r| r.index == row_idx ? new_row : r }
        else
          # Row doesn't exist, create it
          new_cell = Cell.new(row_index: row_idx, column_index: col_idx, value: value, style_index: style_index, formula: formula)
          new_row = Row.new(index: row_idx, cells: [new_cell])
          new_rows = (rows + [new_row]).sort_by!(&:index)
        end
        with(rows: new_rows)
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
        new_unmapped = changes.key?(:unmapped_data) ? changes[:unmapped_data] : unmapped_data
        new_errors = changes.key?(:errors) ? changes[:errors] : errors

        self.class.new(
          name: new_name,
          rows: new_rows,
          columns: new_cols,
          charts: new_charts,
          unmapped_data: new_unmapped,
          errors: new_errors
        )
      end

      # Support pattern matching.
      #: (Array[Symbol]?) -> Hash[Symbol, untyped]
      def deconstruct_keys(_keys)
        { name: name, rows: rows, columns: columns, charts: charts, unmapped_data: unmapped_data, errors: errors }
      end

      # Compare worksheets for equality.
      #: (untyped other) -> bool
      def ==(other)
        return false unless other.is_a?(Worksheet)

        name == other.name && rows == other.rows && columns == other.columns && charts == other.charts
      end
      alias eql? ==

      #: () -> Integer
      def hash
        [self.class, name, rows, columns, charts].hash
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
          errs << "worksheet name cannot contain \\, /, ?, *, [, or ]" if name.match?(%r{[\\/?*\[\]]})
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
    end
  end
end
