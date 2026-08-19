# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Represents a worksheet being streamed sequentially from an XLSX file.
  # Provides O(1) constant-memory streaming over rows and cells.
  #
  # Call {#load} (or {#to_worksheet}) to convert this streaming sheet into an
  # in-memory {Elements::Worksheet} supporting coordinate random access (`sheet["A1"]`).
  #
  # @example Iterate rows and cells in streaming mode (O(1) memory)
  #   Xlsxrb.read("large_data.xlsx") do |sheet|
  #     puts "Processing sheet: #{sheet.name}"
  #     sheet.each_row do |row|
  #       row.each_cell do |cell|
  #         puts "#{cell.ref}: #{cell.value}"
  #       end
  #     end
  #   end
  #
  # @example Load into an in-memory Worksheet for coordinate random access
  #   wb = Xlsxrb.read("data.xlsx")
  #   doc_sheet = wb.sheets.first.load
  #   puts doc_sheet["A1"].value
  #
  # @api public
  class StreamSheet
    [Enumerable].each { |m| include m }

    # @return [String] The sheet name.
    # @api public
    #: String
    attr_reader :name

    # Initializes a streaming worksheet context.
    #
    # @param name [String] The sheet name.
    # @param sheet_xml [String] Raw XML content of the worksheet.
    # @param shared_strings [Array<String>] Shared strings table.
    # @param styles [Hash, nil] Optional parsed styles hash.
    #: (String name, String sheet_xml, Array[String] shared_strings, ?Hash[untyped, untyped]? styles) -> void
    def initialize(name, sheet_xml, shared_strings, styles = nil)
      @name = name
      @sheet_xml = sheet_xml
      @shared_strings = shared_strings
      @styles = styles
    end

    # Iterates over rows in this streaming worksheet with O(1) memory.
    #
    # @overload each_row(&block)
    #   @yield [row]
    #   @yieldparam row [StreamRow, Elements::Row] The current row.
    #   @return [void]
    #
    # @overload each_row
    #   @return [Enumerator<StreamRow | Elements::Row, void>]
    #
    # @api public
    #: () { (StreamRow | Elements::Row) -> void } -> void
    #: () -> Enumerator[StreamRow | Elements::Row, void]
    def each_row
      return enum_for(:each_row) unless block_given?

      Ooxml::WorksheetParser.each_row(@sheet_xml, shared_strings: @shared_strings) do |row|
        if row.is_a?(Elements::Row) || row.is_a?(StreamRow)
          yield row
        else
          yield Xlsxrb.send(:build_row_from_raw, row)
        end
      end
    end

    # Iterates over all cells across all rows continuously with O(1) memory.
    #
    # @overload each_cell(&block)
    #   @yield [cell]
    #   @yieldparam cell [Elements::Cell] The current cell.
    #   @return [void]
    #
    # @overload each_cell
    #   @return [Enumerator<Elements::Cell, void>]
    #
    # @api public
    #: () { (Elements::Cell) -> void } -> void
    #: () -> Enumerator[Elements::Cell, void]
    def each_cell(&)
      return enum_for(:each_cell) unless block_given?

      each_row do |row|
        row.each_cell(&)
      end
    end

    # Default Enumerable iteration delegates to {#each_row}.
    #
    # @overload each(&block)
    #   @yield [row]
    #   @yieldparam row [StreamRow, Elements::Row]
    #   @return [void]
    #
    # @overload each
    #   @return [Enumerator<StreamRow | Elements::Row, void>]
    #
    # @api public
    #: () { (StreamRow | Elements::Row) -> void } -> void
    #: () -> Enumerator[StreamRow | Elements::Row, void]
    def each(&)
      each_row(&)
    end

    # Loads this sheet completely into an in-memory {Elements::Worksheet},
    # enabling coordinate random access (`sheet["A1"]`), row lookups (`row_at`),
    # and immutable cell updates (`update_cell`).
    #
    # @return [Elements::Worksheet] The fully parsed in-memory worksheet.
    # @api public
    #: () -> Elements::Worksheet
    def load
      Xlsxrb.send(:build_worksheet, @name, @sheet_xml, @shared_strings, @styles)
    end
    alias to_worksheet load
  end
end
