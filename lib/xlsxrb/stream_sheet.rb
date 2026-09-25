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
    # @param sheet_source [String, Proc, IO, nil] Raw XML content or chunk supplier/stream.
    # @param shared_strings [Array<String>] Shared strings table.
    # @param styles [Hash, nil] Optional parsed styles hash.
    # @param zip_reader [Ooxml::ZipReader, nil] Optional ZipReader context.
    # @param entry_name [String, nil] Archive entry name for this sheet.
    #: (String name, untyped sheet_source, Array[String] shared_strings, ?Hash[untyped, untyped]? styles, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?) -> void
    def initialize(name, sheet_source, shared_strings, styles = nil, zip_reader: nil, entry_name: nil)
      @name = name
      @sheet_source = sheet_source
      @shared_strings = shared_strings
      @styles = styles
      @zip_reader = zip_reader
      @entry_name = entry_name
      @sheet_xml = sheet_source if sheet_source.is_a?(String)
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
    def each_row(&)
      return enum_for(:each_row) unless block_given?

      source = if @sheet_source
                 @sheet_source
               elsif @zip_reader && @entry_name
                 ->(&blk) { @zip_reader.each_entry_chunk(@entry_name, &blk) }
               end

      Ooxml::WorksheetParser.each_row(source, shared_strings: @shared_strings, &)
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
      Xlsxrb.send(:build_worksheet, @name, raw_sheet_xml, @shared_strings, @styles)
    end
    alias to_worksheet load

    # Returns merged cell ranges (e.g. ["A1:B2"]) for this worksheet.
    #
    # @return [Array<String>]
    # @api public
    #: () -> Array[String]
    def merged_cells
      @merged_cells ||= begin
        listener = Ooxml::Reader::MergeCellsListener.new
        Ooxml::XmlParser.parse(raw_sheet_xml, listener)
        listener.ranges
      end
    end

    # Returns the auto-filter range (e.g. "A1:E100") for this worksheet, or nil if none.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def auto_filter
      @auto_filter ||= begin
        listener = Ooxml::Reader::AutoFilterListener.new
        Ooxml::XmlParser.parse(raw_sheet_xml, listener)
        listener.ref
      end
    end

    # Returns data validation rules configured for this worksheet.
    #
    # @return [Array<Hash[Symbol, untyped]>]
    # @api public
    #: () -> Array[Hash[Symbol, untyped]]
    def data_validations
      @data_validations ||= begin
        listener = Ooxml::Reader::DataValidationsListener.new
        Ooxml::XmlParser.parse(raw_sheet_xml, listener)
        listener.validations
      end
    end

    # Returns conditional formatting rules configured for this worksheet.
    #
    # @return [Array<Hash[Symbol, untyped]>]
    # @api public
    #: () -> Array[Hash[Symbol, untyped]]
    def conditional_formats
      @conditional_formats ||= begin
        listener = Ooxml::Reader::ConditionalFormattingListener.new
        Ooxml::XmlParser.parse(raw_sheet_xml, listener)
        listener.rules
      end
    end

    # Closes any underlying reader resources.
    #: () -> void
    def close
      @zip_reader&.close
      nil
    end

    private

    # Returns the full raw XML string for this worksheet, loaded on demand.
    #: () -> String
    def raw_sheet_xml
      @raw_sheet_xml ||= if @sheet_source.is_a?(String)
                           @sheet_source
                         elsif @zip_reader && @entry_name
                           @zip_reader.read_entry(@entry_name) || ""
                         elsif @sheet_source.respond_to?(:call)
                           buf = +""
                           @sheet_source.call { |c| buf << c }
                           buf
                         else
                           ""
                         end
    end
  end
end
