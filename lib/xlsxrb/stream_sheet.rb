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

    # @return [Elements::Styles, Hash, nil] The parsed styles definition.
    # @api public
    #: (Elements::Styles | Hash[untyped, untyped])?
    attr_reader :styles

    # Initializes a streaming worksheet context.
    #
    # @param name [String] The sheet name.
    # @param sheet_source [String, Proc, IO, nil] Raw XML content or chunk supplier/stream.
    # @param shared_strings [Array<String>] Shared strings table.
    # @param styles [Elements::Styles, Hash, nil] Optional parsed styles.
    # @param zip_reader [Ooxml::ZipReader, nil] Optional ZipReader context.
    # @param entry_name [String, nil] Archive entry name for this sheet.
    # @param date1904 [Boolean] Whether the 1904 date system is active.
    # @param dimension [String, nil] Sheet dimension reference string (e.g. "A1:Z50000").
    # @param trim_empty_rows [Boolean] Whether to omit trailing empty rows during iteration.
    # @param pad_empty_rows [Boolean] Whether to yield empty rows for skipped row numbers.
    # @param pad_empty_cells [Boolean] Whether to pad missing cells within rows.
    # @param images [Array<Elements::Image>, nil] Optional preloaded images array.
    # @param defined_names [Array<Hash, Elements::DefinedName>, nil] Optional workbook defined names.
    # @param sheet_index [Integer, nil] 0-based worksheet index.
    # @param charts [Array<Elements::Chart>, nil] Optional preloaded charts.
    # @param sparkline_groups [Array<Elements::SparklineGroup>, nil] Optional preloaded sparkline groups.
    # @param print_area [String, nil] Optional print area range.
    # @param print_titles [Elements::PrintTitles, nil] Optional repeating print titles.
    # @param header_footer_images [Array<Elements::HeaderFooterImage>, nil] Optional preloaded header/footer images.
    #: (String name, untyped sheet_source, Array[String] shared_strings, ?(Elements::Styles | Hash[untyped, untyped])? styles, ?zip_reader: Ooxml::ZipReader?, ?entry_name: String?, ?state: Symbol, ?date1904: bool, ?dimension: String?, ?trim_empty_rows: bool, ?pad_empty_rows: bool, ?pad_empty_cells: bool, ?images: Array[Elements::Image]?, ?defined_names: Array[untyped]?, ?sheet_index: Integer?, ?charts: Array[Elements::Chart]?, ?sparkline_groups: Array[Elements::SparklineGroup]?, ?print_area: String?, ?print_titles: Elements::PrintTitles?, ?header_footer_images: Array[Elements::HeaderFooterImage]?) -> void
    def initialize(name, sheet_source, shared_strings, styles = nil, zip_reader: nil, entry_name: nil, state: :visible, date1904: false, dimension: nil, trim_empty_rows: false, pad_empty_rows: false, pad_empty_cells: false, images: nil, defined_names: nil, sheet_index: nil, charts: nil, sparkline_groups: nil, print_area: nil, print_titles: nil, header_footer_images: nil)
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
      @trim_empty_rows = trim_empty_rows ? true : false
      @pad_empty_rows = pad_empty_rows ? true : false
      @pad_empty_cells = pad_empty_cells ? true : false
      @images = images
      @defined_names = defined_names || []
      @sheet_index = sheet_index
      @charts = charts
      @sparkline_groups = sparkline_groups
      @print_area = print_area
      @print_titles = print_titles
      @header_footer_images = header_footer_images
    end

    # Returns whether trailing empty rows are omitted during iteration.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def trim_empty_rows?
      @trim_empty_rows ? true : false
    end

    # Returns whether skipped rows are padded during iteration.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def pad_empty_rows?
      @pad_empty_rows ? true : false
    end

    # Returns whether missing cells within rows are padded during iteration.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def pad_empty_cells?
      @pad_empty_cells ? true : false
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

    # Returns the 0-based range of rows defined in the sheet dimension, or nil.
    #
    # @example
    #   sheet.row_range #=> 0..9
    #
    # @return [Range<Integer>, nil]
    # @api public
    #: () -> Range[Integer]?
    def row_range
      bounds = dimension_bounds
      return nil unless bounds

      (bounds[:min_row]..bounds[:max_row])
    end

    # Returns the 0-based range of columns defined in the sheet dimension, or nil.
    #
    # @example
    #   sheet.column_range #=> 0..4
    #
    # @return [Range<Integer>, nil]
    # @api public
    #: () -> Range[Integer]?
    def column_range
      bounds = dimension_bounds
      return nil unless bounds

      (bounds[:min_column]..bounds[:max_column])
    end

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

    # Returns whether the sheet is very hidden (:very_hidden).
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def very_hidden?
      @state == :very_hidden
    end

    # Iterates over rows in this streaming worksheet with O(1) memory.
    #
    # @overload each_row(&block)
    #   @yield [row]
    #   @yieldparam row [StreamRow, Elements::Row] The current row.
    # @overload each_row(trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil, &block)
    #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows (defaults to sheet setting).
    #   @param pad_empty_rows [Boolean, nil] Whether to yield empty rows for skipped row numbers (defaults to sheet setting).
    #   @param pad_empty_cells [Boolean, nil] Whether to pad missing cells within rows (defaults to sheet setting).
    #   @yield [row]
    #   @yieldparam row [StreamRow, Elements::Row] The current row.
    #   @return [void]
    #
    # @overload each_row(trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil)
    #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows (defaults to sheet setting).
    #   @param pad_empty_rows [Boolean, nil] Whether to yield empty rows for skipped row numbers (defaults to sheet setting).
    #   @param pad_empty_cells [Boolean, nil] Whether to pad missing cells within rows (defaults to sheet setting).
    #   @return [Enumerator<StreamRow | Elements::Row, void>]
    #
    # @api public
    #: (?trim_empty_rows: bool?, ?pad_empty_rows: bool?, ?pad_empty_cells: bool?) { (StreamRow | Elements::Row) -> void } -> void
    #: (?trim_empty_rows: bool?, ?pad_empty_rows: bool?, ?pad_empty_cells: bool?) -> Enumerator[StreamRow | Elements::Row, void]
    def each_row(trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil, &)
      return enum_for(:each_row, trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells) unless block_given?

      trim = if trim_empty_rows.nil?
               @trim_empty_rows
             else
               (trim_empty_rows ? true : false)
             end
      pad_rows = if pad_empty_rows.nil?
                   @pad_empty_rows
                 else
                   (pad_empty_rows ? true : false)
                 end
      pad_cells = if pad_empty_cells.nil?
                    @pad_empty_cells
                  else
                    (pad_empty_cells ? true : false)
                  end

      source = if @sheet_source
                 @sheet_source
               elsif @zip_reader && @entry_name
                 ->(&blk) { @zip_reader.each_entry_chunk(@entry_name, &blk) }
               end

      Ooxml::WorksheetParser.each_row(source, shared_strings: @shared_strings, styles: @styles, date1904: @date1904, trim_empty_rows: trim, pad_empty_rows: pad_rows, pad_empty_cells: pad_cells, &)
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
    # @overload each_row_values(type_cast: false, trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil, &block)
    #   @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
    #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows (defaults to sheet setting).
    #   @param pad_empty_rows [Boolean, nil] Whether to yield empty rows for skipped row numbers (defaults to sheet setting).
    #   @param pad_empty_cells [Boolean, nil] Whether to pad missing cells within rows (defaults to sheet setting).
    #   @yield [values]
    #   @yieldparam values [Array<Object>] Row values array.
    #   @return [void]
    #
    # @overload each_row_values(type_cast: false, trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil)
    #   @param type_cast [Boolean] Whether to coerce date/time serial numbers into Date/Time instances.
    #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows (defaults to sheet setting).
    #   @param pad_empty_rows [Boolean, nil] Whether to yield empty rows for skipped row numbers (defaults to sheet setting).
    #   @param pad_empty_cells [Boolean, nil] Whether to pad missing cells within rows (defaults to sheet setting).
    #   @return [Enumerator<Array<Object>, void>]
    #
    # @api public
    #: (?type_cast: bool, ?trim_empty_rows: bool?, ?pad_empty_rows: bool?, ?pad_empty_cells: bool?) { (Array[untyped]) -> void } -> void
    #: (?type_cast: bool, ?trim_empty_rows: bool?, ?pad_empty_rows: bool?, ?pad_empty_cells: bool?) -> Enumerator[Array[untyped], void]
    def each_row_values(type_cast: false, trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil, &block)
      return enum_for(:each_row_values, type_cast: type_cast, trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells) unless block

      each_row(trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells) do |row|
        block.call(row.values(type_cast: type_cast))
      end
    end

    # Default Enumerable iteration delegates to {#each_row}.
    #
    # @overload each(trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil, &block)
    #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows.
    #   @param pad_empty_rows [Boolean, nil] Whether to yield empty rows for skipped row numbers.
    #   @param pad_empty_cells [Boolean, nil] Whether to pad missing cells within rows.
    #   @yield [row]
    #   @yieldparam row [StreamRow, Elements::Row]
    #   @return [void]
    #
    # @overload each(trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil)
    #   @param trim_empty_rows [Boolean, nil] Whether to omit trailing empty rows.
    #   @param pad_empty_rows [Boolean, nil] Whether to yield empty rows for skipped row numbers.
    #   @param pad_empty_cells [Boolean, nil] Whether to pad missing cells within rows.
    #   @return [Enumerator<StreamRow | Elements::Row, void>]
    #
    # @api public
    #: (?trim_empty_rows: bool?, ?pad_empty_rows: bool?, ?pad_empty_cells: bool?) { (StreamRow | Elements::Row) -> void } -> void
    #: (?trim_empty_rows: bool?, ?pad_empty_rows: bool?, ?pad_empty_cells: bool?) -> Enumerator[StreamRow | Elements::Row, void]
    def each(trim_empty_rows: nil, pad_empty_rows: nil, pad_empty_cells: nil, &)
      each_row(trim_empty_rows: trim_empty_rows, pad_empty_rows: pad_empty_rows, pad_empty_cells: pad_empty_cells, &)
    end

    # Loads this sheet completely into an in-memory {Elements::Worksheet},
    # enabling coordinate random access (`sheet["A1"]`) and row lookups (`row_at`).
    #
    # @return [Elements::Worksheet] The fully parsed in-memory worksheet.
    # @api public
    #: () -> Elements::Worksheet
    def load
      Xlsxrb.send(
        :build_worksheet,
        @name,
        raw_sheet_xml,
        @shared_strings,
        @styles,
        state: @state,
        zip_reader: @zip_reader,
        entry_name: @entry_name,
        hyperlinks: hyperlinks,
        comments: comments,
        threaded_comments: threaded_comments,
        images: images,
        charts: charts,
        sparkline_groups: sparkline_groups,
        print_area: print_area,
        print_titles: print_titles,
        date1904: @date1904,
        dimension: dimension,
        trim_empty_rows: @trim_empty_rows,
        pad_empty_rows: @pad_empty_rows,
        pad_empty_cells: @pad_empty_cells,
        columns: columns,
        merged_cells: merged_cells,
        auto_filter: auto_filter,
        filter_columns: filter_columns,
        sort_state: sort_state,
        tables: tables,
        pivot_tables: pivot_tables,
        shapes: shapes,
        freeze_pane: freeze_pane,
        split_pane: split_pane,
        page_setup: page_setup,
        page_margins: page_margins,
        header_footer: header_footer,
        print_options: print_options,
        sheet_protection: sheet_protection,
        sheet_view: sheet_view,
        sheet_properties: sheet_properties,
        sheet_format: sheet_format,
        row_breaks: row_breaks,
        column_breaks: column_breaks,
        protected_ranges: protected_ranges,
        cell_watches: cell_watches,
        ignored_errors: ignored_errors,
        phonetic_properties: phonetic_properties,
        selection: selection,
        sheet_calc_properties: sheet_calc_properties,
        background_image: background_image,
        header_footer_images: header_footer_images
      )
    end

    # Returns hyperlinks configured for this worksheet as a Hash of cell references to {Elements::Hyperlink} objects.
    #
    # @return [Hash<String, Elements::Hyperlink>]
    # @api public
    #: () -> Hash[String, Elements::Hyperlink]
    def hyperlinks
      @hyperlinks ||= Xlsxrb.send(:resolve_hyperlinks, raw_sheet_xml, zip_reader: @zip_reader, entry_name: @entry_name)
    end

    # Returns hyperlink metadata for a specific cell reference, or nil.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, String, Symbol, nil] Optional 0-based column index or column letter/Symbol.
    # @return [Elements::Hyperlink, nil]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Elements::Hyperlink?
    def hyperlink(ref_or_row, col = nil)
      ref = if col
              c_idx = Xlsxrb.column_index(col)
              "#{Xlsxrb.column_letter(c_idx)}#{ref_or_row.to_i + 1}"
            else
              ref_or_row.to_s.upcase
            end
      hyperlinks[ref]
    end

    # Returns comments configured for this worksheet as an Array of {Elements::Comment} objects.
    #
    # @return [Array<Elements::Comment>]
    # @api public
    #: () -> Array[Elements::Comment]
    def comments
      @comments ||= begin
        raw = Xlsxrb.send(:resolve_comments, zip_reader: @zip_reader, entry_name: @entry_name)
        raw.map do |c|
          if c.is_a?(Elements::Comment)
            c
          else
            Elements::Comment.new(
              ref: c[:ref] || c[:cell],
              author: c[:author],
              text: c[:text],
              visible: c[:visible] || false
            )
          end
        end.freeze
      end
    end

    # Returns comment metadata for a specific cell reference, or nil.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, String, Symbol, nil] Optional 0-based column index or column letter/Symbol.
    # @return [Elements::Comment, nil]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Elements::Comment?
    def comment(ref_or_row, col = nil)
      ref = if col
              c_idx = Xlsxrb.column_index(col)
              "#{Xlsxrb.column_letter(c_idx)}#{ref_or_row.to_i + 1}"
            else
              ref_or_row.to_s.upcase
            end
      comments_by_ref[ref]
    end

    # Returns comments indexed by cell reference.
    #
    # @return [Hash<String, Elements::Comment>]
    # @api public
    #: () -> Hash[String, Elements::Comment]
    def comments_by_ref
      @comments_by_ref ||= comments.each_with_object({}) do |c, acc|
        r = c.respond_to?(:ref) ? c.ref : (c[:ref] || c[:cell])
        acc[r.to_s.upcase] = c if r
      end.freeze
    end

    # Returns threaded comments configured for this worksheet as an Array of {Elements::ThreadedComment} objects.
    #
    # @return [Array<Elements::ThreadedComment>]
    # @api public
    #: () -> Array[Elements::ThreadedComment]
    def threaded_comments
      @threaded_comments ||= begin
        raw = Xlsxrb.send(:resolve_threaded_comments, zip_reader: @zip_reader, entry_name: @entry_name)
        raw.map do |tc|
          if tc.is_a?(Elements::ThreadedComment)
            tc
          else
            Elements::ThreadedComment.new(
              id: tc[:id],
              ref: tc[:ref] || tc[:cell],
              author: tc[:author],
              text: tc[:text],
              parent_id: tc[:parent_id],
              resolved: tc[:resolved] || tc[:done] || false,
              created_at: tc[:created_at] || tc[:dt],
              author_id: tc[:author_id] || tc[:person_id]
            )
          end
        end.freeze
      end
    end

    # Returns the root threaded comment for a specific cell reference, or nil.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, String, Symbol, nil] Optional 0-based column index or column letter/Symbol.
    # @return [Elements::ThreadedComment, nil]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Elements::ThreadedComment?
    def threaded_comment(ref_or_row, col = nil)
      ref = if col
              c_idx = Xlsxrb.column_index(col)
              "#{Xlsxrb.column_letter(c_idx)}#{ref_or_row.to_i + 1}"
            else
              ref_or_row.to_s.upcase
            end
      threaded_comments_by_ref[ref]&.first
    end

    # Returns threaded comments for a specific cell reference, or empty array.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, String, Symbol, nil] Optional 0-based column index or column letter/Symbol.
    # @return [Array<Elements::ThreadedComment>]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Array[Elements::ThreadedComment]
    def threaded_comments_at(ref_or_row, col = nil)
      ref = if col
              c_idx = Xlsxrb.column_index(col)
              "#{Xlsxrb.column_letter(c_idx)}#{ref_or_row.to_i + 1}"
            else
              ref_or_row.to_s.upcase
            end
      threaded_comments_by_ref[ref] || []
    end

    # Returns threaded comments indexed by cell reference.
    #
    # @return [Hash<String, Array<Elements::ThreadedComment>>]
    # @api public
    #: () -> Hash[String, Array[Elements::ThreadedComment]]
    def threaded_comments_by_ref
      @threaded_comments_by_ref ||= threaded_comments.each_with_object({}) do |c, acc|
        r = c.ref
        (acc[r.to_s.upcase] ||= []) << c if r
      end.freeze
    end

    # Returns all embedded images in this worksheet.
    #
    # @return [Array<Elements::Image>]
    # @api public
    #: () -> Array[Elements::Image]
    def images
      @images ||= Xlsxrb.send(:resolve_images, zip_reader: @zip_reader, entry_name: @entry_name)
    end

    # Returns all images anchored at a specific cell reference or 0-based coordinate.
    #
    # @param ref_or_row [String, Symbol, Integer] Cell reference (e.g. "A1") or 0-based row index.
    # @param col [Integer, String, Symbol, nil] Optional 0-based column index or column letter/Symbol.
    # @return [Array<Elements::Image>]
    # @api public
    #: (String | Symbol | Integer ref_or_row, ?Integer | String | Symbol | nil col) -> Array[Elements::Image]
    def images_at(ref_or_row, col = nil)
      ref = if col
              c_idx = Xlsxrb.column_index(col)
              "#{Xlsxrb.column_letter(c_idx)}#{ref_or_row.to_i + 1}"
            else
              ref_or_row.to_s.upcase
            end
      images.select { |img| img.ref == ref }
    end

    # Returns merged cell ranges (e.g. ["A1:B2"]) for this worksheet.
    #
    # @return [Array<String>]
    # @api public
    #: () -> Array[String]
    def merged_cells
      @merged_cells ||= if raw_sheet_xml.include?("<mergeCell") || raw_sheet_xml.include?(":mergeCell")
                          listener = Ooxml::Reader::MergeCellsListener.new
                          Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                          listener.ranges
                        else
                          [].freeze
                        end
    end

    # Returns the auto-filter range (e.g. "A1:E100") for this worksheet, or nil if none.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def auto_filter
      @auto_filter ||= auto_filter_data[:ref]
    end

    # Returns column filter rules defined for this worksheet.
    #
    # @return [Hash<Integer, Elements::FilterColumn>]
    # @api public
    #: () -> Hash[Integer, Elements::FilterColumn]
    def filter_columns
      @filter_columns ||= (auto_filter_data[:filter_columns] || {}).to_h do |col_id, filter|
        fc = if filter.is_a?(Elements::FilterColumn)
               filter
             else
               Elements::FilterColumn.from_spec(col_id, filter)
             end
        [col_id, fc]
      end.freeze
    end

    # Returns data validation rules configured for this worksheet.
    #
    # @return [Array<Elements::DataValidation>]
    # @api public
    #: () -> Array[Elements::DataValidation]
    def data_validations
      @data_validations ||= if raw_sheet_xml.include?("<dataValidation") || raw_sheet_xml.include?(":dataValidation")
                              listener = Ooxml::Reader::DataValidationsListener.new
                              Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                              listener.validations.map do |dv|
                                dv.is_a?(Elements::DataValidation) ? dv : Elements::DataValidation.new(**dv)
                              end.freeze
                            else
                              [].freeze
                            end
    end

    # Returns conditional formatting rules configured for this worksheet.
    #
    # Corresponds to ECMA-376 Part 4 §18.3.1.18 (<conditionalFormatting> and <cfRule> elements).
    #
    # @return [Array<Elements::ConditionalFormat>]
    # @api public
    #: () -> Array[Elements::ConditionalFormat]
    def conditional_formats
      @conditional_formats ||= if raw_sheet_xml.include?("<conditionalFormatting") || raw_sheet_xml.include?(":conditionalFormatting")
                                 listener = Ooxml::Reader::ConditionalFormattingListener.new
                                 Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                                 listener.rules.map do |cf|
                                   if cf.is_a?(Elements::ConditionalFormat)
                                     cf
                                   else
                                     rules = (cf[:rules] || []).map { |r| r.is_a?(Elements::ConditionalFormatRule) ? r : Elements::ConditionalFormatRule.new(**r) }
                                     Elements::ConditionalFormat.new(**cf, rules: rules)
                                   end
                                 end.freeze
                               else
                                 [].freeze
                               end
    end

    # Returns tables defined in this worksheet.
    #
    # @return [Array<Elements::Table>]
    # @api public
    #: () -> Array[Elements::Table]
    def tables
      @tables ||= begin
        raw = Xlsxrb.send(:resolve_tables, zip_reader: @zip_reader, entry_name: @entry_name)
        raw.map do |tb|
          if tb.is_a?(Elements::Table)
            tb
          else
            cols = (tb[:columns] || []).map { |col| col.is_a?(Elements::TableColumn) ? col : Elements::TableColumn.new(**col) }
            si = tb[:style_info]
            si = Elements::TableStyleInfo.new(**si) if si.is_a?(Hash)
            Elements::Table.new(**tb, columns: cols, style_info: si)
          end
        end.freeze
      end
    end

    # Returns pivot tables defined in this worksheet.
    #
    # @return [Array<Elements::PivotTable>]
    # @api public
    #: () -> Array[Elements::PivotTable]
    def pivot_tables
      @pivot_tables ||= begin
        raw = Xlsxrb.send(:resolve_pivot_tables, zip_reader: @zip_reader, entry_name: @entry_name)
        raw.map do |pt|
          pt.is_a?(Elements::PivotTable) ? pt : Elements::PivotTable.new(**pt)
        end.freeze
      end
    end

    # Returns shapes defined in this worksheet.
    #
    # @return [Array<Elements::Shape>]
    # @api public
    #: () -> Array[Elements::Shape]
    def shapes
      @shapes ||= begin
        raw = Xlsxrb.send(:resolve_shapes, zip_reader: @zip_reader, entry_name: @entry_name)
        raw.map do |s|
          s.is_a?(Elements::Shape) ? s : Elements::Shape.new(**s)
        end.freeze
      end
    end

    # Returns freeze pane settings for this worksheet, or nil if none.
    #
    # @return [Elements::FreezePane, nil]
    # @api public
    #: () -> Elements::FreezePane?
    def freeze_pane
      pane = sheet_view_data[:pane]
      return unless pane && pane[:state] == :frozen

      pane.is_a?(Elements::FreezePane) ? pane : Elements::FreezePane.new(**pane)
    end

    # Returns split pane settings for this worksheet, or nil if none.
    #
    # @return [Elements::SplitPane, nil]
    # @api public
    #: () -> Elements::SplitPane?
    def split_pane
      pane = sheet_view_data[:pane]
      return unless pane && pane[:state] != :frozen

      pane.is_a?(Elements::SplitPane) ? pane : Elements::SplitPane.new(**pane)
    end

    # Returns sheet view properties for this worksheet.
    #
    # @return [Elements::SheetView]
    # @api public
    #: () -> Elements::SheetView
    def sheet_view
      v = sheet_view_data[:view] || {}
      v.is_a?(Elements::SheetView) ? v : Elements::SheetView.new(**v)
    end

    # Returns sort state configuration for this worksheet, or nil if none.
    #
    # @return [Elements::SortState, nil]
    # @api public
    #: () -> Elements::SortState?
    def sort_state
      @sort_state ||= if raw_sheet_xml.include?("<sortState")
                        listener = Ooxml::Reader::SortStateListener.new
                        Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                        st = listener.sort_state
                        if st
                          st.is_a?(Elements::SortState) ? st : Elements::SortState.new(**st)
                        end
                      end
    end

    # Returns selection settings for this worksheet, or nil if none.
    #
    # @return [Elements::Selection, nil]
    # @api public
    #: () -> Elements::Selection?
    def selection
      sel = sheet_view_data[:selection]
      return unless sel

      sel.is_a?(Elements::Selection) ? sel : Elements::Selection.new(**sel)
    end

    # Returns sheet protection settings as a hash, or nil if unprotected.
    #
    # @return [Elements::SheetProtection, nil]
    # @api public
    #: () -> Elements::SheetProtection?
    def sheet_protection
      @sheet_protection ||= if raw_sheet_xml.include?("<sheetProtection")
                              listener = Ooxml::Reader::SheetProtectionListener.new
                              Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                              sp = listener.protection
                              if sp
                                sp.is_a?(Elements::SheetProtection) ? sp : Elements::SheetProtection.new(**sp)
                              end
                            end
    end

    # Returns sheet-level properties (tabColor, outlinePr, sheetPr) for this worksheet.
    #
    # @return [Elements::SheetProperties]
    # @api public
    #: () -> Elements::SheetProperties
    def sheet_properties
      @sheet_properties ||= begin
        raw = if raw_sheet_xml.include?("<sheetPr") || raw_sheet_xml.include?("<sheetCalcPr")
                listener = Ooxml::Reader::SheetPropertiesListener.new
                Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                listener.properties
              else
                {}
              end
        raw.is_a?(Elements::SheetProperties) ? raw : Elements::SheetProperties.new(**raw)
      end
    end

    # Returns the internal VBA or application code name of the worksheet, or nil if not specified.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def code_name
      sheet_properties.code_name
    end

    # Returns whether fit-to-page is enabled for page setup on this worksheet.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def fit_to_page?
      sheet_properties.fit_to_page?
    end

    # Returns whether full calculation on load is enabled on this worksheet.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def full_calc_on_load?
      sheet_calc_properties&.to_h&.[](:full_calc_on_load) == true || sheet_properties.to_h[:full_calc_on_load] == true
    end

    # Returns sheet calculation properties (<sheetCalcPr>), or nil if not present.
    #
    # @return [Elements::SheetCalcProperties, nil]
    # @api public
    #: () -> Elements::SheetCalcProperties?
    def sheet_calc_properties
      sheet_properties.full_calc_on_load? ? Elements::SheetCalcProperties.new(full_calc_on_load: true) : nil
    end

    # Returns the worksheet tab color specification, or nil if not set.
    #
    # @return [String, Hash, nil] RGB hex string (e.g. "FFFF0000") or color hash (e.g. { theme: 3, tint: -0.5 }).
    # @api public
    #: () -> (String | Hash[Symbol, untyped])?
    def tab_color
      props = sheet_properties
      props = props.to_h if props.is_a?(Elements::MetadataRecord)
      return props[:tab_color] if props[:tab_color]
      return { theme: props[:tab_color_theme], tint: props[:tab_color_tint] } if props[:tab_color_theme]
      return { indexed: props[:tab_color_indexed] } if props[:tab_color_indexed]
      return { auto: true } if props[:tab_color_auto]

      nil
    end

    # Returns the outline and grouping properties of the worksheet, or an empty hash if none configured.
    #
    # @return [Hash]
    # @api public
    #: () -> Hash[Symbol, untyped]
    def outline_properties
      props = sheet_properties
      props = props.to_h if props.is_a?(Elements::MetadataRecord)
      return props[:outline_properties] if props[:outline_properties]

      res = {}
      res[:summary_below] = props[:summary_below] unless props[:summary_below].nil?
      res[:summary_right] = props[:summary_right] unless props[:summary_right].nil?
      res[:show_outline_symbols] = props[:show_outline_symbols] unless props[:show_outline_symbols].nil?
      res[:apply_styles] = props[:apply_styles] unless props[:apply_styles].nil?
      res
    end

    # Returns sheet format properties (defaultRowHeight, defaultColWidth, baseColWidth) for this worksheet.
    #
    # @return [Elements::SheetFormat]
    # @api public
    #: () -> Elements::SheetFormat
    def sheet_format
      @sheet_format ||= begin
        raw = if raw_sheet_xml.include?("<sheetFormatPr")
                listener = Ooxml::Reader::SheetFormatListener.new
                Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                listener.properties
              else
                {}
              end
        raw.is_a?(Elements::SheetFormat) ? raw : Elements::SheetFormat.new(**raw)
      end
    end

    # Returns page setup configurations for this worksheet.
    #
    # @return [Elements::PageSetup]
    # @api public
    #: () -> Elements::PageSetup
    def page_setup
      @page_setup ||= begin
        ps = print_page_data[:page_setup] || {}
        ps.is_a?(Elements::PageSetup) ? ps : Elements::PageSetup.new(**ps)
      end
    end

    # Returns page margins for this worksheet, or nil if none.
    #
    # @return [Elements::PageMargins, nil]
    # @api public
    #: () -> Elements::PageMargins?
    def page_margins
      @page_margins ||= begin
        pm = print_page_data[:page_margins]
        if pm
          pm.is_a?(Elements::PageMargins) ? pm : Elements::PageMargins.new(**pm)
        end
      end
    end

    # Returns header/footer configurations for this worksheet.
    #
    # @return [Elements::HeaderFooter]
    # @api public
    #: () -> Elements::HeaderFooter
    def header_footer
      @header_footer ||= begin
        hf = print_page_data[:header_footer] || {}
        hf.is_a?(Elements::HeaderFooter) ? hf : Elements::HeaderFooter.new(**hf)
      end
    end

    # Returns print options for this worksheet.
    #
    # @return [Elements::PrintOptions]
    # @api public
    #: () -> Elements::PrintOptions
    def print_options
      @print_options ||= begin
        po = print_page_data[:print_options] || {}
        po.is_a?(Elements::PrintOptions) ? po : Elements::PrintOptions.new(**po)
      end
    end

    # Returns explicit row page breaks for this worksheet.
    #
    # @return [Array<Elements::PageBreak>]
    # @api public
    #: () -> Array[Elements::PageBreak]
    def row_breaks
      @row_breaks ||= (print_page_data[:row_breaks] || []).map do |rb|
        rb.is_a?(Elements::PageBreak) ? rb : Elements::PageBreak.new(**rb)
      end.freeze
    end

    # Returns explicit column page breaks for this worksheet.
    #
    # @return [Array<Elements::PageBreak>]
    # @api public
    #: () -> Array[Elements::PageBreak]
    def column_breaks
      @column_breaks ||= (print_page_data[:column_breaks] || []).map do |cb|
        cb.is_a?(Elements::PageBreak) ? cb : Elements::PageBreak.new(**cb)
      end.freeze
    end

    # Returns protected ranges defined for this worksheet.
    #
    # @return [Array<Elements::ProtectedRange>]
    # @api public
    #: () -> Array[Elements::ProtectedRange]
    def protected_ranges
      @protected_ranges ||= if raw_sheet_xml.include?("<protectedRanges")
                              listener = Ooxml::Reader::ProtectedRangesListener.new
                              Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                              listener.ranges.map do |pr|
                                pr.is_a?(Elements::ProtectedRange) ? pr : Elements::ProtectedRange.new(**pr)
                              end.freeze
                            else
                              [].freeze
                            end
    end

    # Returns cell watch cell references defined for this worksheet.
    #
    # @return [Array<String>]
    # @api public
    #: () -> Array[String]
    def cell_watches
      @cell_watches ||= if raw_sheet_xml.include?("<cellWatches")
                          listener = Ooxml::Reader::CellWatchesListener.new
                          Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                          listener.watches
                        else
                          []
                        end
    end

    # Returns ignored error rules defined for this worksheet.
    #
    # @return [Array<Elements::IgnoredError>]
    # @api public
    #: () -> Array[Elements::IgnoredError]
    def ignored_errors
      @ignored_errors ||= if raw_sheet_xml.include?("<ignoredErrors")
                            listener = Ooxml::Reader::IgnoredErrorsListener.new
                            Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                            listener.errors.map do |ie|
                              ie.is_a?(Elements::IgnoredError) ? ie : Elements::IgnoredError.new(**ie)
                            end.freeze
                          else
                            [].freeze
                          end
    end

    # Returns phonetic properties defined for this worksheet, or nil.
    #
    # @return [Elements::PhoneticProperties, nil]
    # @api public
    #: () -> Elements::PhoneticProperties?
    def phonetic_properties
      @phonetic_properties ||= if raw_sheet_xml.include?("<phoneticPr")
                                 listener = Ooxml::Reader::PhoneticPrListener.new
                                 Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                                 pp = listener.phonetic_pr
                                 if pp
                                   pp.is_a?(Elements::PhoneticProperties) ? pp : Elements::PhoneticProperties.new(**pp)
                                 end
                               end
    end

    # Returns the background image configured for this worksheet, or nil.
    #
    # @return [Elements::BackgroundImage, nil]
    # @api public
    #: () -> Elements::BackgroundImage?
    def background_image
      @background_image ||= begin
        bi = Xlsxrb.send(:resolve_background_image, raw_sheet_xml, zip_reader: @zip_reader, entry_name: @entry_name)
        if bi
          bi.is_a?(Elements::BackgroundImage) ? bi : Elements::BackgroundImage.new(**bi)
        end
      end
    end

    # Closes any underlying reader resources.
    #: () -> void
    def close
      @zip_reader&.close
      nil
    end

    # Returns 0-based bounds { min_row: r1, max_row: r2, min_column: c1, max_column: c2 } from dimension ref.
    #
    # @example
    #   sheet.dimension_bounds #=> { min_row: 0, max_row: 4, min_column: 0, max_column: 2 }
    #
    # @return [Hash{Symbol => Integer}, nil]
    # @api public
    #: () -> Hash[Symbol, Integer]?
    def dimension_bounds
      return @dimension_bounds if defined?(@dimension_bounds)

      ref = dimension
      return @dimension_bounds = nil unless ref

      parts = ref.split(":", 2)
      first_coords = Elements::Cell._parse_ref(parts[0])
      return @dimension_bounds = nil unless first_coords

      last_coords = parts[1] ? Elements::Cell._parse_ref(parts[1]) : first_coords
      return @dimension_bounds = nil unless last_coords

      @dimension_bounds = {
        min_row: first_coords[0],
        max_row: last_coords[0],
        min_column: first_coords[1],
        max_column: last_coords[1]
      }.freeze
    end

    # Returns column definitions for this worksheet.
    #
    # @return [Array<Elements::Column>]
    # @api public
    #: () -> Array[Elements::Column]
    def columns
      @columns ||= begin
        raw_columns = Ooxml::WorksheetParser.parse_columns(raw_sheet_xml)
        raw_columns.map do |rc|
          col_unmapped = {}
          col_unmapped[:style_index] = rc[:style_index] if rc[:style_index]
          col_unmapped[:min] = rc[:min] if rc[:min]
          col_unmapped[:max] = rc[:max] if rc[:max]
          Elements::Column.new(
            index: (rc[:min] || 1) - 1,
            width: rc[:width],
            hidden: rc[:hidden] || false,
            custom_width: rc[:custom_width] || false,
            outline_level: rc[:outline_level],
            style_index: rc[:style_index],
            unmapped_data: col_unmapped
          )
        end
      end
    end

    # Returns charts defined in this worksheet.
    #
    # @return [Array<Elements::Chart>]
    # @api public
    #: () -> Array[Elements::Chart]
    def charts
      @charts ||= Xlsxrb.send(:resolve_charts, zip_reader: @zip_reader, entry_name: @entry_name).freeze
    end

    # Returns sparkline groups defined in this worksheet.
    #
    # @return [Array<Elements::SparklineGroup>]
    # @api public
    #: () -> Array[Elements::SparklineGroup]
    def sparkline_groups
      @sparkline_groups ||= begin
        groups = if raw_sheet_xml&.include?("sparklineGroup")
                   Xlsxrb.send(:resolve_sparkline_groups, raw_sheet_xml)
                 else
                   []
                 end
        groups.map { |sg| sg.is_a?(Elements::SparklineGroup) ? sg : Elements::SparklineGroup.new(**sg) }.freeze
      end
    end

    # Returns the print area range for this worksheet, or nil if unset.
    #
    # @return [String, nil]
    # @api public
    #: () -> String?
    def print_area
      return @print_area if defined?(@print_area_resolved) && @print_area_resolved

      @print_area_resolved = true
      @print_area ||= Xlsxrb.send(:extract_sheet_print_area, @defined_names, @name, @sheet_index)
    end

    # Returns print titles configuration for this worksheet, or nil if unset.
    #
    # @return [Elements::PrintTitles, nil]
    # @api public
    #: () -> Elements::PrintTitles?
    def print_titles
      return @print_titles if defined?(@print_titles_resolved) && @print_titles_resolved

      @print_titles_resolved = true
      @print_titles ||= Xlsxrb.send(:extract_sheet_print_titles, @defined_names, @name, @sheet_index)
    end

    # Returns whether the worksheet has any structural or schema validation errors.
    #
    # @return [Boolean]
    # @api public
    #: () -> bool
    def valid?
      errors.empty?
    end

    # Returns validation error messages for this worksheet.
    #
    # @return [Array<String>]
    # @api public
    #: () -> Array[String]
    def errors
      []
    end

    # Returns unmapped metadata for this worksheet.
    #
    # @return [Hash[Symbol, untyped]]
    # @api public
    #: () -> Hash[Symbol, untyped]
    def unmapped_data
      {}
    end

    # Returns header/footer images configured for this worksheet.
    #
    # @return [Array<Elements::HeaderFooterImage>]
    # @api public
    #: () -> Array[Elements::HeaderFooterImage]
    def header_footer_images
      @header_footer_images ||= begin
        imgs = if raw_sheet_xml&.include?("legacyDrawingHF")
                 Xlsxrb.send(:resolve_header_footer_images, raw_sheet_xml, zip_reader: @zip_reader, entry_name: @entry_name)
               else
                 []
               end
        imgs.map { |img| img.is_a?(Elements::HeaderFooterImage) ? img : Elements::HeaderFooterImage.new(**img) }.freeze
      end
    end

    private

    #: () -> Hash[Symbol, untyped]
    def auto_filter_data
      @auto_filter_data ||= if raw_sheet_xml.include?("<autoFilter")
                              listener = Ooxml::Reader::AutoFilterListener.new
                              Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                              { ref: listener.ref, filter_columns: listener.filter_columns }
                            else
                              { ref: nil, filter_columns: {} }
                            end
    end

    #: () -> Hash[Symbol, untyped]
    def sheet_view_data
      @sheet_view_data ||= if raw_sheet_xml.include?("<sheetView")
                             listener = Ooxml::Reader::SheetViewListener.new
                             Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                             { view: listener.view, pane: listener.pane, selection: listener.selection }
                           else
                             { view: {}, pane: nil, selection: nil }
                           end
    end

    #: () -> Hash[Symbol, untyped]
    def print_page_data
      @print_page_data ||= if raw_sheet_xml.match?(/<(?:pageSetup|pageMargins|headerFooter|printOptions|rowBreaks|colBreaks)\b/)
                             listener = Ooxml::Reader::PrintPageListener.new
                             Ooxml::XmlParser.parse(raw_sheet_xml, listener)
                             {
                               page_margins: listener.page_margins,
                               page_setup: listener.page_setup,
                               header_footer: listener.header_footer,
                               print_options: listener.print_options,
                               row_breaks: listener.row_breaks,
                               column_breaks: listener.column_breaks
                             }
                           else
                             {
                               page_margins: nil,
                               page_setup: {},
                               header_footer: {},
                               print_options: {},
                               row_breaks: [],
                               column_breaks: []
                             }
                           end
    end

    #: () -> String
    def raw_sheet_xml
      @raw_sheet_xml ||= if @sheet_source.is_a?(String)
                           @sheet_source
                         elsif @zip_reader && @entry_name
                           @zip_reader.read_entry(@entry_name) || ""
                         elsif @sheet_source.respond_to?(:call)
                           buf = +""
                           @sheet_source.call { |c| buf << c }
                           buf.force_encoding(Encoding::UTF_8)
                         else
                           ""
                         end
    end

    # Extracts dimension reference from XML source lazily without parsing the entire file.
    #: () -> String?
    def extract_dimension
      if @sheet_source.is_a?(String)
        @sheet_source.slice(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/, 1)
      elsif @zip_reader && @entry_name
        buf = +""
        dim = nil
        @zip_reader.each_entry_chunk(@entry_name) do |chunk|
          buf << chunk
          if (m = buf.match(/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*\bref=["']([^"']+)["']/))
            dim = m[1]
            break
          end
          break if buf.include?("<sheetData") || buf.include?(":sheetData")
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
