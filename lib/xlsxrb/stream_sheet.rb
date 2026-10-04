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

    # @return [Symbol] The sheet visibility state (:visible, :hidden, or :very_hidden).
    # @api public
    #: Symbol
    attr_reader :state

    # @return [Hash, nil] The parsed styles definition hash.
    # @api public
    #: Hash[untyped, untyped]?
    attr_reader :styles

    # Initializes a streaming worksheet context.
    #
    # @param name [String] The sheet name.
    # @param sheet_source [String, Proc, IO, nil] Raw XML content or chunk supplier/stream.
    # @param shared_strings [Array<String>] Shared strings table.
    # @param styles [Hash, nil] Optional parsed styles hash.
    # @param zip_reader [Ooxml::ZipReader, nil] Optional ZipReader context.
    # @param entry_name [String, nil] Archive entry name for this sheet.
    # @param date1904 [Boolean] Whether the 1904 date system is active.
    # @param dimension [String, nil] Sheet dimension reference string (e.g. "A1:Z50000").
    #: (String name, untyped sheet_source, Array[String] shared_strings, ?Hash[untyped, untyped]? styles, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?, ?state: Symbol, ?date1904: bool, ?dimension: String?) -> void
    def initialize(name, sheet_source, shared_strings, styles = nil, zip_reader: nil, entry_name: nil, state: :visible, date1904: false, dimension: nil)
      @name = name
      @sheet_source = sheet_source
      @shared_strings = shared_strings
      @styles = styles
      @zip_reader = zip_reader
      @entry_name = entry_name
      @state = state ? state.to_sym : :visible
      @date1904 = date1904 ? true : false
      @dimension = dimension
      @dimension_checked = !dimension.nil?
      @sheet_xml = sheet_source if sheet_source.is_a?(String)
    end

    # Returns the dimension reference string (e.g. "A1:Z50000"), or nil if not present.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def dimension
      return @dimension if @dimension_checked

      @dimension_checked = true
      @dimension = extract_dimension
    end

    # Returns the 1-based index of the first row defined in the sheet dimension, or nil.
    #
    # @return [Integer, nil]
    # @api public
    #: () -> Integer?
    def first_row
      dimension_bounds&.[](0)
    end

    # Returns the 1-based index of the first column defined in the sheet dimension, or nil.
    #
    # @return [Integer, nil]
    # @api public
    #: () -> Integer?
    def first_column
      dimension_bounds&.[](1)
    end
    alias first_col first_column

    # Returns the 1-based index of the last row defined in the sheet dimension, or nil.
    #
    # @return [Integer, nil]
    # @api public
    #: () -> Integer?
    def last_row
      dimension_bounds&.[](2)
    end

    # Returns the 1-based index of the last column defined in the sheet dimension, or nil.
    #
    # @return [Integer, nil]
    # @api public
    #: () -> Integer?
    def last_column
      dimension_bounds&.[](3)
    end
    alias last_col last_column

    # Returns whether the sheet uses the 1904 date system.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def date1904?
      @date1904
    end

    # Returns whether the sheet is hidden (:hidden or :very_hidden).
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def hidden?
      @state == :hidden || @state == :very_hidden
    end

    # Returns whether the sheet is visible.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def visible?
      @state == :visible
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

      Ooxml::WorksheetParser.each_row(source, shared_strings: @shared_strings, styles: @styles, date1904: @date1904, &)
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

    # Iterates over row values directly as Arrays without wrapping cells in Elements::Cell objects.
    #
    # @overload each_row_values(type_cast: false, &block)
    #   @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
    #   @yield [values]
    #   @yieldparam values [Array<Object>] Row values array.
    #   @return [void]
    #
    # @overload each_row_values(type_cast: false)
    #   @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
    #   @return [Enumerator<Array<Object>, void>]
    #
    # @api public
    #: (?type_cast: bool) { (Array[untyped]) -> void } -> void
    #: (?type_cast: bool) -> Enumerator[Array[untyped], void]
    def each_row_values(type_cast: false, &block)
      return enum_for(:each_row_values, type_cast: type_cast) unless block

      each_row do |row|
        block.call(row.values(type_cast: type_cast))
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
      Xlsxrb.send(:build_worksheet, @name, raw_sheet_xml, @shared_strings, @styles, state: @state, zip_reader: @zip_reader, entry_name: @entry_name, hyperlinks: hyperlinks, comments: comments, date1904: @date1904, dimension: dimension)
    end
    alias to_worksheet load

    # Returns hyperlinks configured for this worksheet as a Hash of cell references to hyperlink hashes.
    #
    # @return [Hash<String, Hash[Symbol, untyped]>]
    # @api public
    #: () -> Hash[String, Hash[Symbol, untyped]]
    def hyperlinks
      @hyperlinks ||= Xlsxrb.send(:resolve_hyperlinks, raw_sheet_xml, zip_reader: @zip_reader, entry_name: @entry_name)
    end

    # Returns hyperlink metadata for a specific cell reference, or nil.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, nil] Optional 0-based column index.
    # @return [Hash, nil]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer? col) -> Hash[Symbol, untyped]?
    def hyperlink(ref_or_row, col = nil)
      ref = if col
              "#{Elements::Cell.column_letter(col)}#{ref_or_row.to_i + 1}"
            else
              ref_or_row.to_s.upcase
            end
      hyperlinks[ref]
    end

    # Returns comments configured for this worksheet as an Array of comment hashes.
    #
    # @return [Array<Hash[Symbol, untyped]>]
    # @api public
    #: () -> Array[Hash[Symbol, untyped]]
    def comments
      @comments ||= Xlsxrb.send(:resolve_comments, zip_reader: @zip_reader, entry_name: @entry_name)
    end

    # Returns comment metadata for a specific cell reference, or nil.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, nil] Optional 0-based column index.
    # @return [Hash, nil]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer? col) -> Hash[Symbol, untyped]?
    def comment(ref_or_row, col = nil)
      ref = if col
              "#{Elements::Cell.column_letter(col)}#{ref_or_row.to_i + 1}"
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
      @comments_by_ref ||= comments.each_with_object({}) do |c, acc|
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
      target_row, target_col = if col
                                 [ref_or_row.to_i, col.to_i]
                               else
                                 parsed = Elements::Cell.parse_ref(ref_or_row.to_s)
                                 return nil unless parsed

                                 parsed
                               end

      each_row do |row|
        next if row.index < target_row
        return row.cell_at(target_col) if row.index == target_row
        break if row.index > target_row
      end
      nil
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

    # Returns 1-based bounds [first_row, first_col, last_row, last_col] from dimension ref.
    #: () -> Array[Integer]?
    def dimension_bounds
      return @dimension_bounds if defined?(@dimension_bounds)

      ref = dimension
      return @dimension_bounds = nil unless ref

      parts = ref.split(":", 2)
      first_coords = Elements::Cell.parse_ref(parts[0])
      return @dimension_bounds = nil unless first_coords

      last_coords = parts[1] ? Elements::Cell.parse_ref(parts[1]) : first_coords
      return @dimension_bounds = nil unless last_coords

      @dimension_bounds = [
        first_coords[0] + 1,
        first_coords[1] + 1,
        last_coords[0] + 1,
        last_coords[1] + 1
      ].freeze
    end

    # Extracts dimension reference from XML source lazily without parsing the entire file.
    #: () -> String?
    def extract_dimension
      if @sheet_source.is_a?(String)
        @sheet_source.slice(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/, 1)
      elsif @zip_reader && @entry_name
        stream = @zip_reader.open_entry_io(@entry_name)
        buf = +""
        dim = nil
        begin
          while (chunk = stream.read(65_536))
            buf << chunk
            if (m = buf.match(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/))
              dim = m[1]
              break
            end
            break if buf.include?("<sheetData") || buf.include?(":sheetData")
          end
        ensure
          stream.close
        end
        dim
      elsif @sheet_source.respond_to?(:call)
        buf = +""
        dim = nil
        @sheet_source.call do |chunk|
          buf << chunk
          if (m = buf.match(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/))
            dim = m[1]
            break
          end
          break if buf.include?("<sheetData") || buf.include?(":sheetData")
        end
        dim
      elsif @sheet_source.respond_to?(:seek) && @sheet_source.respond_to?(:pos)
        pos = @sheet_source.pos
        buf = @sheet_source.read(65_536) || ""
        @sheet_source.seek(pos, IO::SEEK_SET)
        buf.slice(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/, 1)
      else
        raw_sheet_xml.slice(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/, 1)
      end
    end
  end
end
