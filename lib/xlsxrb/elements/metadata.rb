# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  # Immutable value objects that describe workbook content and metadata.
  module Elements
    # Rejects keyword arguments that a value object does not have.
    #
    # Prepended to the singleton class of each value object, so Klass.new(**opts) raises
    # ArgumentError for an unknown key instead of dropping it. A class can declare
    # KEY_ALIASES, mapping old key names to the member they were renamed to: an alias is
    # dropped when its member is also given (readers emit both), and otherwise raises with
    # the new name.
    #
    # @api private
    module StrictKeywords
      #: (*untyped args, **untyped kwargs) -> untyped
      def new(*args, **kwargs)
        return super if kwargs.empty?

        allowed = members
        return super if kwargs.each_key.all? { |key| allowed.include?(key) }

        super(*args, **StrictKeywords.normalize(self, kwargs))
      end

      # Removes alias keys whose member is also given; raises for any other unknown key.
      #
      # @param klass [Class] The value object class.
      # @param kwargs [Hash{Symbol => Object}] Keyword arguments given to new.
      # @return [Hash{Symbol => Object}]
      # @raise [ArgumentError] for an unknown key or an alias given without its member.
      #: (Class klass, Hash[Symbol, untyped] kwargs) -> Hash[Symbol, untyped]
      def self.normalize(klass, kwargs)
        params = klass.instance_method(:initialize).parameters
        # Value objects with an options member collect other keys there through **extra.
        return kwargs if params.any? { |type, _| type == :keyrest }

        allowed = klass.members | params.filter_map { |type, name| name if %i[key keyreq].include?(type) }
        aliases = klass.const_defined?(:KEY_ALIASES, false) ? klass.const_get(:KEY_ALIASES, false) : {}
        result = kwargs.dup
        kwargs.each_key do |key|
          next if allowed.include?(key)

          member = aliases[key]
          if member && !kwargs[member].nil?
            result.delete(key)
            next
          end

          message = "unknown keyword: #{key.inspect} for #{klass.name}"
          if member
            message << "; use #{member.inspect}"
          elsif defined?(DidYouMean::SpellChecker) && (suggestion = DidYouMean::SpellChecker.new(dictionary: allowed.map(&:to_s)).correct(key.to_s).first)
            message << "; did you mean #{suggestion.to_sym.inspect}?"
          end
          raise ArgumentError, message
        end
        result
      end
    end

    # Hyperlink is defined in types.rb, before this file, and does not include MetadataRecord.
    Hyperlink.singleton_class.prepend(StrictKeywords)

    # Shared mixin for metadata Data value objects.
    #
    # @api public
    module MetadataRecord
      #: (Module base) -> void
      def self.included(base)
        base.singleton_class.prepend(StrictKeywords)
      end

      # Returns whether all attributes are nil or empty.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty?
        to_h.values.all? { |v| v.nil? || (v.respond_to?(:empty?) && v.empty?) }
      end

      # Converts the value object to a Hash of non-nil attributes.
      #
      # @return [Hash{Symbol => Object}]
      # @api public
      #: () -> Hash[Symbol, untyped]
      def to_h
        res = super.compact.transform_values do |val|
          if val.is_a?(Array)
            # Nested Arrays, such as the [name, count] pairs of AppProperties#heading_pairs,
            # respond to to_h but are not key-value pairs, so they stay as they are.
            val.map { |v| v.respond_to?(:to_h) && !v.is_a?(String) && !v.is_a?(Array) ? v.to_h : v }
          elsif val.respond_to?(:to_h) && !val.is_a?(String) && !val.is_a?(Hash) && !val.is_a?(RichText)
            # RichText is Enumerable over its runs, so Enumerable#to_h would treat each run as a
            # key-value pair. It stays a value object, which the writers accept as cell and comment text.
            val.to_h
          else
            val
          end
        end
        if members.include?(:options) && res[:options].is_a?(Hash)
          opts = res.delete(:options)
          opts.each { |ok, ov| res[ok] ||= ov }
        end
        res
      end

      alias to_hash to_h
    end

    # Core document properties (docProps/core.xml).
    # @api public
    CoreProperties = Data.define(
      :title, :subject, :creator, :keywords, :description, :last_modified_by,
      :revision, :created, :modified, :category, :content_status,
      :language
    ) do
      include MetadataRecord

      #: (?title: String?, ?subject: String?, ?creator: String?, ?keywords: String?, ?description: String?, ?last_modified_by: String?, ?revision: String?, ?created: untyped, ?modified: untyped, ?category: String?, ?content_status: String?, ?language: String?) -> void
      def initialize(title: nil, subject: nil, creator: nil, keywords: nil, description: nil, last_modified_by: nil, revision: nil, created: nil, modified: nil, category: nil, content_status: nil, language: nil)
        super(**{
          title: title, subject: subject, creator: creator, keywords: keywords, description: description,
          last_modified_by: last_modified_by, revision: revision, created: created, modified: modified,
          category: category, content_status: content_status, language: language
        })
      end
    end

    # Application document properties (docProps/app.xml).
    # @api public
    AppProperties = Data.define(
      :app_name, :app_version, :doc_security, :scale_crop, :heading_pairs,
      :titles_of_parts, :manager, :company, :links_up_to_date, :shared_doc,
      :hyperlink_base, :hyperlinks_changed
    ) do
      include MetadataRecord

      private :scale_crop, :links_up_to_date, :shared_doc, :hyperlinks_changed

      #: (?app_name: String?, ?app_version: String?, ?doc_security: Integer?, ?scale_crop: bool?, ?heading_pairs: Array[untyped]?, ?titles_of_parts: Array[untyped]?, ?manager: String?, ?company: String?, ?links_up_to_date: bool?, ?shared_doc: bool?, ?hyperlink_base: String?, ?hyperlinks_changed: bool?) -> void
      def initialize(app_name: nil, app_version: nil, doc_security: nil, scale_crop: nil, heading_pairs: nil, titles_of_parts: nil, manager: nil, company: nil, links_up_to_date: nil, shared_doc: nil, hyperlink_base: nil, hyperlinks_changed: nil)
        super(**{
          app_name: app_name, app_version: app_version, doc_security: doc_security, scale_crop: scale_crop,
          heading_pairs: heading_pairs, titles_of_parts: titles_of_parts, manager: manager, company: company,
          links_up_to_date: links_up_to_date, shared_doc: shared_doc, hyperlink_base: hyperlink_base,
          hyperlinks_changed: hyperlinks_changed
        })
      end

      # Returns whether scale and crop are enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def scale_crop?
        scale_crop ? true : false
      end

      # Returns whether links are up to date.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def links_up_to_date?
        links_up_to_date ? true : false
      end

      # Returns whether the document is shared.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def shared_doc?
        shared_doc ? true : false
      end

      # Returns whether hyperlinks have changed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hyperlinks_changed?
        hyperlinks_changed ? true : false
      end
    end

    # Workbook-level properties (<workbookPr> and <workbookView>).
    # @api public
    WorkbookProperties = Data.define(
      :date1904, :code_name, :active_sheet, :first_sheet,
      :show_sheet_tabs, :tab_ratio, :full_calc_on_load, :filter_privacy,
      :default_theme_version, :refresh_all,
      :update_links
    ) do
      include MetadataRecord

      private :date1904, :show_sheet_tabs, :full_calc_on_load, :filter_privacy, :refresh_all

      #: (?date1904: bool?, ?code_name: String?, ?active_sheet: Integer?, ?first_sheet: Integer?, ?show_sheet_tabs: bool?, ?tab_ratio: Integer?, ?full_calc_on_load: bool?, ?filter_privacy: bool?, ?default_theme_version: String?, ?refresh_all: bool?, ?update_links: String?) -> void
      def initialize(date1904: nil, code_name: nil, active_sheet: nil, first_sheet: nil, show_sheet_tabs: nil, tab_ratio: nil, full_calc_on_load: nil, filter_privacy: nil, default_theme_version: nil, refresh_all: nil, update_links: nil)
        super(**{
          date1904: date1904, code_name: code_name, active_sheet: active_sheet, first_sheet: first_sheet,
          show_sheet_tabs: show_sheet_tabs, tab_ratio: tab_ratio, full_calc_on_load: full_calc_on_load,
          filter_privacy: filter_privacy, default_theme_version: default_theme_version, refresh_all: refresh_all, update_links: update_links
        })
      end

      # Returns whether the 1904 date system is used.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date1904?
        date1904 ? true : false
      end

      # Returns whether filter privacy is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def filter_privacy?
        filter_privacy ? true : false
      end

      # Returns whether sheet tabs are visible.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_sheet_tabs?
        show_sheet_tabs != false
      end

      # Returns whether full calculation is performed on load.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def full_calc_on_load?
        full_calc_on_load ? true : false
      end

      # Returns whether all external data ranges and PivotTables are refreshed on open.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def refresh_all?
        refresh_all ? true : false
      end
    end

    # Workbook protection settings (<workbookProtection>).
    # @api public
    WorkbookProtection = Data.define(
      :workbook_password, :lock_structure, :lock_windows, :revisions_password,
      :lock_revision, :algorithm_name, :hash_value, :salt_value, :spin_count,
      :revisions_algorithm_name, :revisions_hash_value, :revisions_salt_value, :revisions_spin_count
    ) do
      include MetadataRecord

      private :lock_structure, :lock_windows, :lock_revision

      #: (?workbook_password: String?, ?lock_structure: bool?, ?lock_windows: bool?, ?revisions_password: String?, ?lock_revision: bool?, ?algorithm_name: String?, ?hash_value: String?, ?salt_value: String?, ?spin_count: Integer?, ?revisions_algorithm_name: String?, ?revisions_hash_value: String?, ?revisions_salt_value: String?, ?revisions_spin_count: Integer?) -> void
      def initialize(workbook_password: nil, lock_structure: nil, lock_windows: nil, revisions_password: nil, lock_revision: nil, algorithm_name: nil, hash_value: nil, salt_value: nil, spin_count: nil, revisions_algorithm_name: nil, revisions_hash_value: nil, revisions_salt_value: nil, revisions_spin_count: nil)
        super(**{
          workbook_password: workbook_password, lock_structure: lock_structure, lock_windows: lock_windows,
          revisions_password: revisions_password, lock_revision: lock_revision, algorithm_name: algorithm_name,
          hash_value: hash_value, salt_value: salt_value, spin_count: spin_count, revisions_algorithm_name: revisions_algorithm_name, revisions_hash_value: revisions_hash_value, revisions_salt_value: revisions_salt_value, revisions_spin_count: revisions_spin_count
        })
      end

      # Returns whether the workbook structure is locked.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def lock_structure?
        lock_structure ? true : false
      end

      # Returns whether workbook windows are locked.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def lock_windows?
        lock_windows ? true : false
      end

      # Returns whether workbook revisions are locked.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def lock_revision?
        lock_revision ? true : false
      end
    end

    # Worksheet page setup properties (<pageSetup>).
    # @api public
    PageSetup = Data.define(
      :paper_size, :orientation, :scale, :fit_to_width, :fit_to_height,
      :first_page_number, :use_first_page_number, :black_and_white, :draft,
      :cell_comments, :use_printer_defaults, :copies, :horizontal_centered,
      :vertical_centered, :id,
      :page_order, :horizontal_dpi, :vertical_dpi
    ) do
      include MetadataRecord

      private :use_first_page_number, :black_and_white, :draft, :use_printer_defaults, :horizontal_centered, :vertical_centered

      #: (?paper_size: (Integer | String)?, ?orientation: (Symbol | String)?, ?scale: Integer?, ?fit_to_width: Integer?, ?fit_to_height: Integer?, ?first_page_number: Integer?, ?use_first_page_number: bool?, ?black_and_white: bool?, ?draft: bool?, ?cell_comments: String?, ?use_printer_defaults: bool?, ?copies: Integer?, ?horizontal_centered: bool?, ?vertical_centered: bool?, ?id: String?, ?page_order: String?, ?horizontal_dpi: Integer?, ?vertical_dpi: Integer?) -> void
      def initialize(paper_size: nil, orientation: nil, scale: nil, fit_to_width: nil, fit_to_height: nil, first_page_number: nil, use_first_page_number: nil, black_and_white: nil, draft: nil, cell_comments: nil, use_printer_defaults: nil, copies: nil, horizontal_centered: nil, vertical_centered: nil, id: nil, page_order: nil, horizontal_dpi: nil, vertical_dpi: nil)
        super(**{
          paper_size: paper_size, orientation: orientation, scale: scale, fit_to_width: fit_to_width,
          fit_to_height: fit_to_height, first_page_number: first_page_number, use_first_page_number: use_first_page_number,
          black_and_white: black_and_white, draft: draft, cell_comments: cell_comments,
          use_printer_defaults: use_printer_defaults, copies: copies, horizontal_centered: horizontal_centered,
          vertical_centered: vertical_centered, id: id, page_order: page_order, horizontal_dpi: horizontal_dpi, vertical_dpi: vertical_dpi
        })
      end

      # Returns whether black and white print mode is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def black_and_white?
        black_and_white ? true : false
      end

      # Returns whether draft print mode is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def draft?
        draft ? true : false
      end

      # Returns whether first page number is used.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def use_first_page_number?
        use_first_page_number ? true : false
      end

      # Returns whether printer defaults are used.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def use_printer_defaults?
        use_printer_defaults ? true : false
      end

      # Returns whether page is horizontally centered.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def horizontal_centered?
        horizontal_centered ? true : false
      end

      # Returns whether page is vertically centered.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def vertical_centered?
        vertical_centered ? true : false
      end
    end

    # Worksheet page margins (<pageMargins>).
    # @api public
    PageMargins = Data.define(:left, :right, :top, :bottom, :header, :footer) do
      include MetadataRecord

      #: (?left: Float?, ?right: Float?, ?top: Float?, ?bottom: Float?, ?header: Float?, ?footer: Float?) -> void
      def initialize(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil)
        super(**{ left: left, right: right, top: top, bottom: bottom, header: header, footer: footer })
      end
    end

    # Worksheet header and footer (<headerFooter>).
    # @api public
    HeaderFooter = Data.define(
      :odd_header, :odd_footer, :even_header, :even_footer,
      :first_header, :first_footer, :different_odd_even, :different_first,
      :scale_with_doc, :align_with_margins
    ) do
      include MetadataRecord

      private :different_odd_even, :different_first, :scale_with_doc, :align_with_margins

      #: (?odd_header: String?, ?odd_footer: String?, ?even_header: String?, ?even_footer: String?, ?first_header: String?, ?first_footer: String?, ?different_odd_even: bool?, ?different_first: bool?, ?scale_with_doc: bool?, ?align_with_margins: bool?) -> void
      def initialize(odd_header: nil, odd_footer: nil, even_header: nil, even_footer: nil, first_header: nil, first_footer: nil, different_odd_even: nil, different_first: nil, scale_with_doc: nil, align_with_margins: nil)
        super(**{
          odd_header: odd_header, odd_footer: odd_footer, even_header: even_header,
          even_footer: even_footer, first_header: first_header, first_footer: first_footer,
          different_odd_even: different_odd_even, different_first: different_first,
          scale_with_doc: scale_with_doc, align_with_margins: align_with_margins
        })
      end

      # Returns whether different first page header/footer is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def different_first?
        different_first ? true : false
      end

      # Returns whether different odd and even page header/footer is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def different_odd_even?
        different_odd_even ? true : false
      end

      # Returns whether header and footer scale with document.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def scale_with_doc?
        scale_with_doc ? true : false
      end

      # Returns whether header and footer align with page margins.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def align_with_margins?
        align_with_margins ? true : false
      end
    end

    # Worksheet print options (<printOptions>).
    # @api public
    PrintOptions = Data.define(:grid_lines, :grid_lines_set, :headings, :horizontal_centered, :vertical_centered) do
      include MetadataRecord

      private :grid_lines, :grid_lines_set, :headings, :horizontal_centered, :vertical_centered

      #: (?grid_lines: bool?, ?grid_lines_set: bool?, ?headings: bool?, ?horizontal_centered: bool?, ?vertical_centered: bool?) -> void
      def initialize(grid_lines: nil, grid_lines_set: nil, headings: nil, horizontal_centered: nil, vertical_centered: nil)
        super(**{
          grid_lines: grid_lines, grid_lines_set: grid_lines_set, headings: headings,
          horizontal_centered: horizontal_centered, vertical_centered: vertical_centered
        })
      end

      # Returns whether grid lines are printed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def grid_lines?
        grid_lines ? true : false
      end

      # Returns whether the grid lines setting was explicitly set.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def grid_lines_set?
        grid_lines_set ? true : false
      end

      # Returns whether row and column headings are printed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def headings?
        headings ? true : false
      end

      # Returns whether printing is centered horizontally.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def horizontal_centered?
        horizontal_centered ? true : false
      end

      # Returns whether printing is centered vertically.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def vertical_centered?
        vertical_centered ? true : false
      end
    end

    # Worksheet repeating print titles.
    # @api public
    PrintTitles = Data.define(:rows, :columns) do
      include MetadataRecord

      #: (?rows: String?, ?columns: String?) -> void
      def initialize(rows: nil, columns: nil)
        super(**{ rows: rows, columns: columns })
      end
    end

    # Worksheet freeze pane settings (<pane>).
    # @api public
    #
    # +top_left_cell+ is the first visible cell of the bottom-right pane (the topLeftCell
    # attribute). When nil, the writer uses the cell just past the frozen rows and columns.
    FreezePane = Data.define(:row, :column, :top_left_cell, :active_pane, :state) do
      include MetadataRecord

      #: (?row: Integer?, ?column: (Integer | String | Symbol)?, ?top_left_cell: String?, ?active_pane: String?, ?state: Symbol?) -> void
      def initialize(row: nil, column: nil, top_left_cell: nil, active_pane: nil, state: :frozen)
        super(**{ row: row, column: column, top_left_cell: top_left_cell, active_pane: active_pane, state: state || :frozen })
      end
    end

    # Worksheet split pane settings (<pane>).
    # @api public
    #
    # +top_left_cell+ is the first visible cell of the bottom-right pane (the topLeftCell attribute).
    SplitPane = Data.define(:x_split, :y_split, :top_left_cell, :active_pane, :state) do
      include MetadataRecord

      #: (?x_split: (Integer | Float)?, ?y_split: (Integer | Float)?, ?top_left_cell: String?, ?active_pane: String?, ?state: String?) -> void
      def initialize(x_split: nil, y_split: nil, top_left_cell: nil, active_pane: nil, state: nil)
        super(**{ x_split: x_split, y_split: y_split, top_left_cell: top_left_cell, active_pane: active_pane, state: state })
      end
    end

    # Worksheet selection / active cell (<selection>).
    # @api public
    Selection = Data.define(:pane, :active_cell, :active_cell_id, :sqref) do
      include MetadataRecord

      #: (?pane: String?, ?active_cell: String?, ?active_cell_id: Integer?, ?sqref: String?) -> void
      def initialize(pane: nil, active_cell: nil, active_cell_id: nil, sqref: nil)
        super(**{ pane: pane, active_cell: active_cell, active_cell_id: active_cell_id, sqref: sqref })
      end
    end

    # Worksheet protection options (<sheetProtection>).
    # @api public
    SheetProtection = Data.define(
      :password, :sheet, :objects, :scenarios, :format_cells, :format_columns,
      :format_rows, :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns,
      :delete_rows, :select_locked_cells, :select_unlocked_cells, :sort, :auto_filter,
      :pivot_tables, :algorithm_name, :hash_value, :salt_value, :spin_count
    ) do
      include MetadataRecord

      private :sheet, :objects, :scenarios, :format_cells, :format_columns, :format_rows, :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns, :delete_rows, :select_locked_cells, :select_unlocked_cells, :sort, :auto_filter, :pivot_tables

      #: (?password: String?, ?sheet: bool?, ?objects: bool?, ?scenarios: bool?, ?format_cells: bool?, ?format_columns: bool?, ?format_rows: bool?, ?insert_columns: bool?, ?insert_rows: bool?, ?insert_hyperlinks: bool?, ?delete_columns: bool?, ?delete_rows: bool?, ?select_locked_cells: bool?, ?select_unlocked_cells: bool?, ?sort: bool?, ?auto_filter: bool?, ?pivot_tables: bool?, ?algorithm_name: String?, ?hash_value: String?, ?salt_value: String?, ?spin_count: Integer?) -> void
      def initialize(password: nil, sheet: nil, objects: nil, scenarios: nil, format_cells: nil, format_columns: nil, format_rows: nil, insert_columns: nil, insert_rows: nil, insert_hyperlinks: nil, delete_columns: nil, delete_rows: nil, select_locked_cells: nil, select_unlocked_cells: nil, sort: nil, auto_filter: nil, pivot_tables: nil, algorithm_name: nil, hash_value: nil, salt_value: nil, spin_count: nil)
        super(**{
          password: password, sheet: sheet, objects: objects, scenarios: scenarios,
          format_cells: format_cells, format_columns: format_columns, format_rows: format_rows,
          insert_columns: insert_columns, insert_rows: insert_rows, insert_hyperlinks: insert_hyperlinks,
          delete_columns: delete_columns, delete_rows: delete_rows, select_locked_cells: select_locked_cells,
          select_unlocked_cells: select_unlocked_cells, sort: sort, auto_filter: auto_filter,
          pivot_tables: pivot_tables, algorithm_name: algorithm_name, hash_value: hash_value,
          salt_value: salt_value, spin_count: spin_count
        })
      end

      # Returns whether worksheet protection is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def sheet?
        sheet ? true : false
      end

      # Returns whether objects are protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def objects?
        objects ? true : false
      end

      # Returns whether scenarios are protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def scenarios?
        scenarios ? true : false
      end

      # Returns whether formatting cells is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def format_cells?
        format_cells ? true : false
      end

      # Returns whether formatting columns is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def format_columns?
        format_columns ? true : false
      end

      # Returns whether formatting rows is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def format_rows?
        format_rows ? true : false
      end

      # Returns whether inserting columns is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def insert_columns?
        insert_columns ? true : false
      end

      # Returns whether inserting rows is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def insert_rows?
        insert_rows ? true : false
      end

      # Returns whether inserting hyperlinks is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def insert_hyperlinks?
        insert_hyperlinks ? true : false
      end

      # Returns whether deleting columns is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def delete_columns?
        delete_columns ? true : false
      end

      # Returns whether deleting rows is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def delete_rows?
        delete_rows ? true : false
      end

      # Returns whether selecting locked cells is allowed/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def select_locked_cells?
        select_locked_cells ? true : false
      end

      # Returns whether selecting unlocked cells is allowed/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def select_unlocked_cells?
        select_unlocked_cells ? true : false
      end

      # Returns whether sorting is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def sort?
        sort ? true : false
      end

      # Returns whether auto filtering is locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def auto_filter?
        auto_filter ? true : false
      end

      # Returns whether pivot tables are locked/protected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def pivot_tables?
        pivot_tables ? true : false
      end
    end

    # Worksheet properties (<sheetPr>).
    # @api public
    SheetProperties = Data.define(
      :tab_color, :filter_mode, :enable_format_conditions_calculation, :page_setup_properties,
      :code_name, :fit_to_page, :full_calc_on_load, :tab_color_theme, :tab_color_tint, :tab_color_indexed,
      :tab_color_auto, :outline_properties, :auto_page_breaks, :sync_horizontal, :sync_vertical, :sync_ref,
      :transition_evaluation, :transition_entry, :published,
      :apply_styles, :summary_below, :summary_right, :show_outline_symbols
    ) do
      include MetadataRecord

      private :filter_mode, :enable_format_conditions_calculation, :fit_to_page, :full_calc_on_load, :tab_color_auto, :auto_page_breaks, :sync_horizontal, :sync_vertical, :transition_evaluation, :transition_entry, :published, :apply_styles, :summary_below, :summary_right, :show_outline_symbols

      #: (?tab_color: String?, ?filter_mode: bool?, ?enable_format_conditions_calculation: bool?, ?page_setup_properties: Hash[Symbol, untyped]?, ?code_name: String?, ?fit_to_page: bool?, ?full_calc_on_load: bool?, ?tab_color_theme: Integer?, ?tab_color_tint: Float?, ?tab_color_indexed: Integer?, ?tab_color_auto: bool?, ?outline_properties: Hash[Symbol, untyped]?, ?auto_page_breaks: bool?, ?sync_horizontal: bool?, ?sync_vertical: bool?, ?sync_ref: String?, ?transition_evaluation: bool?, ?transition_entry: bool?, ?published: bool?, ?apply_styles: bool?, ?summary_below: bool?, ?summary_right: bool?, ?show_outline_symbols: bool?) -> void
      def initialize(tab_color: nil, filter_mode: nil, enable_format_conditions_calculation: nil, page_setup_properties: nil, code_name: nil, fit_to_page: nil, full_calc_on_load: nil, tab_color_theme: nil, tab_color_tint: nil, tab_color_indexed: nil, tab_color_auto: nil, outline_properties: nil, auto_page_breaks: nil, sync_horizontal: nil, sync_vertical: nil, sync_ref: nil, transition_evaluation: nil, transition_entry: nil, published: nil, apply_styles: nil, summary_below: nil, summary_right: nil, show_outline_symbols: nil)
        psp = page_setup_properties
        ftp = if fit_to_page.nil?
                psp.is_a?(Hash) ? psp[:fit_to_page] : nil
              else
                fit_to_page
              end
        apb = if auto_page_breaks.nil?
                psp.is_a?(Hash) ? psp[:auto_page_breaks] : nil
              else
                auto_page_breaks
              end
        out_pr = outline_properties || {}
        app_st = apply_styles.nil? ? out_pr[:apply_styles] : apply_styles
        sum_b = summary_below.nil? ? out_pr[:summary_below] : summary_below
        sum_r = summary_right.nil? ? out_pr[:summary_right] : summary_right
        show_sym = show_outline_symbols.nil? ? out_pr[:show_outline_symbols] : show_outline_symbols
        super(**{
          tab_color: tab_color, filter_mode: filter_mode,
          enable_format_conditions_calculation: enable_format_conditions_calculation,
          page_setup_properties: psp,
          code_name: code_name, fit_to_page: ftp, full_calc_on_load: full_calc_on_load,
          tab_color_theme: tab_color_theme, tab_color_tint: tab_color_tint, tab_color_indexed: tab_color_indexed,
          tab_color_auto: tab_color_auto, outline_properties: outline_properties, auto_page_breaks: apb,
          sync_horizontal: sync_horizontal, sync_vertical: sync_vertical, sync_ref: sync_ref,
          transition_evaluation: transition_evaluation, transition_entry: transition_entry, published: published,
          apply_styles: app_st, summary_below: sum_b, summary_right: sum_r, show_outline_symbols: show_sym
        })
      end

      # Returns whether filter mode is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def filter_mode?
        filter_mode ? true : false
      end

      # Returns whether format conditions calculation is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def enable_format_conditions_calculation?
        enable_format_conditions_calculation ? true : false
      end

      # Returns whether full calc on load is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def full_calc_on_load?
        full_calc_on_load ? true : false
      end

      # Returns whether tab color is set to automatic.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def tab_color_auto?
        tab_color_auto ? true : false
      end

      # Returns whether automatic page breaks are enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def auto_page_breaks?
        auto_page_breaks ? true : false
      end

      # Returns whether horizontal scrolling is synced.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def sync_horizontal?
        sync_horizontal ? true : false
      end

      # Returns whether vertical scrolling is synced.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def sync_vertical?
        sync_vertical ? true : false
      end

      # Returns whether transition evaluation is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def transition_evaluation?
        transition_evaluation ? true : false
      end

      # Returns whether transition entry is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def transition_entry?
        transition_entry ? true : false
      end

      # Returns whether sheet is published.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def published?
        published ? true : false
      end

      # Returns whether fit to page is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def fit_to_page?
        fit_to_page ? true : false
      end

      # Returns whether outline styles are applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_styles?
        apply_styles ? true : false
      end

      # Returns whether summary rows appear below detail rows.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def summary_below?
        summary_below ? true : false
      end

      # Returns whether summary columns appear to the right of detail columns.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def summary_right?
        summary_right ? true : false
      end

      # Returns whether outline symbols are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_outline_symbols?
        show_outline_symbols ? true : false
      end
    end

    # Worksheet format properties (<sheetFormatPr>).
    # @api public
    SheetFormat = Data.define(
      :default_row_height, :default_column_width, :custom_height, :zero_height,
      :thick_top, :thick_bottom, :outline_level_row, :outline_level_column, :base_column_width
    ) do
      include MetadataRecord

      private :custom_height, :zero_height, :thick_top, :thick_bottom

      #: (?default_row_height: (Float | Integer)?, ?default_column_width: (Float | Integer)?, ?custom_height: bool?, ?zero_height: bool?, ?thick_top: bool?, ?thick_bottom: bool?, ?outline_level_row: Integer?, ?outline_level_column: Integer?, ?base_column_width: (Float | Integer)?) -> void
      def initialize(default_row_height: nil, default_column_width: nil, custom_height: nil, zero_height: nil, thick_top: nil, thick_bottom: nil, outline_level_row: nil, outline_level_column: nil, base_column_width: nil)
        super(**{
          default_row_height: default_row_height, default_column_width: default_column_width,
          custom_height: custom_height, zero_height: zero_height, thick_top: thick_top,
          thick_bottom: thick_bottom, outline_level_row: outline_level_row,
          outline_level_column: outline_level_column, base_column_width: base_column_width
        })
      end

      # Returns whether custom row height is set.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def custom_height?
        custom_height ? true : false
      end

      # Returns whether zero row height is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def zero_height?
        zero_height ? true : false
      end

      # Returns whether thick top border is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def thick_top?
        thick_top ? true : false
      end

      # Returns whether thick bottom border is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def thick_bottom?
        thick_bottom ? true : false
      end
    end

    # Worksheet view configuration (<sheetView>).
    # @api public
    SheetView = Data.define(
      :workbook_view_id, :show_grid_lines, :show_row_col_headers, :show_ruler,
      :show_zeros, :right_to_left, :tab_selected, :zoom_scale, :zoom_scale_normal,
      :zoom_scale_page_layout_view, :zoom_scale_sheet_layout_view, :view,
      :window_protection, :show_formulas, :show_outline_symbols, :default_grid_color, :show_white_space,
      :top_left_cell, :color_id
    ) do
      include MetadataRecord

      private :show_grid_lines, :show_row_col_headers, :show_ruler, :show_zeros, :right_to_left, :tab_selected, :window_protection, :show_formulas, :show_outline_symbols, :default_grid_color, :show_white_space

      #: (?workbook_view_id: Integer?, ?show_grid_lines: bool?, ?show_row_col_headers: bool?, ?show_ruler: bool?, ?show_zeros: bool?, ?right_to_left: bool?, ?tab_selected: bool?, ?zoom_scale: Integer?, ?zoom_scale_normal: Integer?, ?zoom_scale_page_layout_view: Integer?, ?zoom_scale_sheet_layout_view: Integer?, ?view: String?, ?window_protection: bool?, ?show_formulas: bool?, ?show_outline_symbols: bool?, ?default_grid_color: bool?, ?show_white_space: bool?, ?top_left_cell: String?, ?color_id: Integer?) -> void
      def initialize(
        workbook_view_id: nil, show_grid_lines: nil, show_row_col_headers: nil, show_ruler: nil,
        show_zeros: nil, right_to_left: nil, tab_selected: nil, zoom_scale: nil, zoom_scale_normal: nil,
        zoom_scale_page_layout_view: nil, zoom_scale_sheet_layout_view: nil, view: nil,
        window_protection: nil, show_formulas: nil, show_outline_symbols: nil, default_grid_color: nil, show_white_space: nil,
        top_left_cell: nil, color_id: nil
      )
        super(**{
          workbook_view_id: workbook_view_id, show_grid_lines: show_grid_lines,
          show_row_col_headers: show_row_col_headers, show_ruler: show_ruler,
          show_zeros: show_zeros, right_to_left: right_to_left, tab_selected: tab_selected,
          zoom_scale: zoom_scale, zoom_scale_normal: zoom_scale_normal,
          zoom_scale_page_layout_view: zoom_scale_page_layout_view,
          zoom_scale_sheet_layout_view: zoom_scale_sheet_layout_view, view: view,
          window_protection: window_protection, show_formulas: show_formulas,
          show_outline_symbols: show_outline_symbols, default_grid_color: default_grid_color,
          show_white_space: show_white_space, top_left_cell: top_left_cell, color_id: color_id
        })
      end

      # Returns whether window protection is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def window_protection?
        window_protection ? true : false
      end

      # Returns whether formulas are displayed instead of calculated values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_formulas?
        show_formulas ? true : false
      end

      # Returns whether grid lines are shown in the view.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_grid_lines?
        show_grid_lines ? true : false
      end

      # Returns whether row and column headers are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_row_col_headers?
        show_row_col_headers ? true : false
      end

      # Returns whether zero values are displayed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_zeros?
        show_zeros ? true : false
      end

      # Returns whether right-to-left layout is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def right_to_left?
        right_to_left ? true : false
      end

      # Returns whether this sheet tab is selected.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def tab_selected?
        tab_selected ? true : false
      end

      # Returns whether the ruler is shown in page layout view.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_ruler?
        show_ruler ? true : false
      end

      # Returns whether outline symbols are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_outline_symbols?
        show_outline_symbols ? true : false
      end

      # Returns whether the default grid line color is used.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def default_grid_color?
        default_grid_color ? true : false
      end

      # Returns whether white space is displayed in page layout view.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_white_space?
        show_white_space ? true : false
      end
    end

    # Phonetic text properties (<phoneticPr>).
    # @api public
    PhoneticProperties = Data.define(:font_id, :type, :alignment) do
      include MetadataRecord

      #: (?font_id: Integer?, ?type: String?, ?alignment: String?) -> void
      def initialize(font_id: nil, type: nil, alignment: nil)
        super(**{ font_id: font_id, type: type, alignment: alignment })
      end
    end

    # Worksheet calculation properties (<sheetCalcPr>).
    # @api public
    SheetCalcProperties = Data.define(:full_calc_on_load) do
      include MetadataRecord

      private :full_calc_on_load

      #: (?full_calc_on_load: bool?) -> void
      def initialize(full_calc_on_load: nil)
        super(**{ full_calc_on_load: full_calc_on_load })
      end

      # Returns whether full calculation is performed on load.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def full_calc_on_load?
        full_calc_on_load ? true : false
      end
    end

    # Sort state configuration (<sortState>).
    # @api public
    SortState = Data.define(:ref, :case_sensitive, :sort_method, :conditions) do
      include MetadataRecord

      private :case_sensitive

      #: (?ref: String?, ?case_sensitive: bool?, ?sort_method: String?, ?conditions: Array[Hash[Symbol, untyped]]?, **untyped) -> void
      def initialize(ref: nil, case_sensitive: nil, sort_method: nil, conditions: nil, **extra)
        conds = conditions || extra[:sort_conditions]
        super(**{ ref: ref, case_sensitive: case_sensitive, sort_method: sort_method, conditions: conds })
      end

      # Returns whether sorting is case sensitive.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def case_sensitive?
        case_sensitive ? true : false
      end
    end

    # Worksheet background image (<picture>).
    # @api public
    BackgroundImage = Data.define(:r_id, :path, :extension, :data, :content_type) do
      include MetadataRecord

      #: (?r_id: String?, ?path: String?, ?extension: String?, ?data: String?, ?content_type: String?) -> void
      def initialize(r_id: nil, path: nil, extension: nil, data: nil, content_type: nil)
        super(**{ r_id: r_id, path: path, extension: extension, data: data, content_type: content_type })
      end
    end

    # Table styles configuration in stylesheet (<tableStyles>).
    # @api public
    TableStyles = Data.define(:default_table_style, :default_pivot_style, :styles) do
      include MetadataRecord

      #: (?default_table_style: String?, ?default_pivot_style: String?, ?styles: Array[Hash[Symbol, untyped]]?) -> void
      def initialize(default_table_style: nil, default_pivot_style: nil, styles: nil)
        super(**{ default_table_style: default_table_style, default_pivot_style: default_pivot_style, styles: styles })
      end
    end

    # Colors palette configuration in stylesheet (<colors>).
    # @api public
    Colors = Data.define(:indexed, :mru) do
      include MetadataRecord

      #: (?indexed: Array[String]?, ?mru: Array[Hash[Symbol, untyped]]?) -> void
      def initialize(indexed: nil, mru: nil)
        super(**{ indexed: indexed, mru: mru })
      end
    end

    # User-defined custom document property (<property>).
    # @api public
    CustomProperty = Data.define(:name, :value, :type) do
      include MetadataRecord

      #: (?name: String?, ?value: untyped, ?type: Symbol?) -> void
      def initialize(name: nil, value: nil, type: nil)
        super(**{ name: name, value: value, type: type })
      end

      # @return [String]
      # @api public
      #: () -> String
      def to_s
        value.to_s
      end

      # Matches against another CustomProperty or scalar value.
      #: (untyped other) -> bool
      def ==(other)
        if other.is_a?(CustomProperty)
          super
        else
          value == other
        end
      end
    end

    # AutoFilter column configuration (<filterColumn>).
    # @api public
    FilterColumn = Data.define(:column_id, :filters, :custom_filters, :dynamic_filter, :top10, :color_filter, :icon_filter, :hidden, :show_button, :blank, :custom_filters_and) do
      include MetadataRecord

      private :hidden, :show_button

      #: (?column_id: Integer?, ?filters: Array[untyped]?, ?custom_filters: Array[untyped]?, ?dynamic_filter: Hash[Symbol, untyped]?, ?top10: Hash[Symbol, untyped]?, ?color_filter: Hash[Symbol, untyped]?, ?icon_filter: Hash[Symbol, untyped]?, ?hidden: bool?, ?show_button: bool?, ?blank: bool?, ?custom_filters_and: bool?) -> void
      def initialize(column_id: nil, filters: nil, custom_filters: nil, dynamic_filter: nil, top10: nil, color_filter: nil, icon_filter: nil, hidden: false, show_button: true, blank: nil, custom_filters_and: nil)
        super(**{ column_id: column_id, filters: filters, custom_filters: custom_filters, dynamic_filter: dynamic_filter, top10: top10, color_filter: color_filter, icon_filter: icon_filter, hidden: hidden || false, show_button: show_button != false, blank: blank, custom_filters_and: custom_filters_and })
      end

      # Returns whether the filter column is hidden.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hidden?
        hidden ? true : false
      end

      # Returns whether the filter dropdown button is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_button?
        show_button ? true : false
      end

      # Builds a FilterColumn from the filter Hash that the reader and Worksheet#filter_column
      # use: { type: :filters, values: [...] }, { type: :custom, operator:, val: } or
      # { type: :custom, filters: [...] }, { type: :dynamic, ... }, { type: :top10, ... },
      # { type: :color_filter, ... }, or { type: :icon_filter, ... }. A Hash with member
      # keys and no :type is passed to new as is.
      #
      # @param column_id [Integer] 0-based column within the filter range.
      # @param spec [Hash]
      # @return [FilterColumn]
      # @raise [ArgumentError] for an unknown :type.
      # @api private
      #: (Integer column_id, Hash[Symbol, untyped] spec) -> FilterColumn
      def self.from_spec(column_id, spec)
        return new(column_id: column_id, **spec.except(:column_id)) unless spec.key?(:type)

        base = { column_id: column_id, hidden: spec[:hidden] || spec[:hidden_button] || false, show_button: spec[:show_button] != false }
        detail = spec.except(:type, :hidden, :hidden_button, :show_button)
        case spec[:type].to_sym
        when :filters then new(**base, filters: Array(detail[:values]), blank: detail[:blank])
        when :custom then new(**base, custom_filters: detail[:filters] || [{ operator: detail[:operator], val: detail[:val] }], custom_filters_and: detail[:and])
        when :dynamic then new(**base, dynamic_filter: detail)
        when :top10 then new(**base, top10: detail)
        when :color_filter then new(**base, color_filter: detail)
        when :icon_filter then new(**base, icon_filter: detail)
        else raise ArgumentError, "unknown filter type: #{spec[:type].inspect}"
        end
      end

      # Returns the filter Hash that the worksheet writer emits (see .from_spec).
      #
      # @return [Hash{Symbol => Object}]
      # @api private
      #: () -> Hash[Symbol, untyped]
      def to_spec
        spec = if (filters && !filters.empty?) || blank then { type: :filters, values: filters || [], blank: blank }.compact
               elsif custom_filters && custom_filters.size == 1 then { type: :custom, operator: custom_filters.first[:operator], val: custom_filters.first[:val] }
               elsif custom_filters && !custom_filters.empty? then { type: :custom, filters: custom_filters, and: custom_filters_and }.compact
               elsif dynamic_filter then { type: :dynamic, **dynamic_filter }
               elsif top10 then { type: :top10, **top10 }
               elsif color_filter then { type: :color_filter, **color_filter }
               elsif icon_filter then { type: :icon_filter, **icon_filter }
               else {}
               end
        spec[:hidden] = true if hidden?
        spec[:show_button] = false unless show_button?
        spec
      end
    end

    # Workbook defined name (<definedName>).
    # @api public
    DefinedName = Data.define(:name, :value, :local_sheet_id, :local_sheet_name, :hidden, :comment) do
      include MetadataRecord

      private :hidden

      #: (?name: String?, ?value: String?, ?local_sheet_id: Integer?, ?local_sheet_name: String?, ?hidden: bool?, ?comment: String?) -> void
      def initialize(name: nil, value: nil, local_sheet_id: nil, local_sheet_name: nil, hidden: nil, comment: nil)
        super(**{ name: name, value: value || "", local_sheet_id: local_sheet_id, local_sheet_name: local_sheet_name, hidden: hidden, comment: comment })
      end

      # Returns whether the defined name is hidden.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hidden?
        hidden ? true : false
      end

      # @return [String]
      # @api public
      #: () -> String
      def to_s
        value.to_s
      end
    end

    # Explicit page row or column break (<brk>).
    # @api public
    PageBreak = Data.define(:id, :max, :min, :man, :pt) do
      include MetadataRecord

      #: (?id: Integer?, ?max: Integer?, ?min: Integer?, ?man: (Integer | bool)?, ?pt: bool?) -> void
      # rubocop:disable Naming/MethodParameterName
      def initialize(id: nil, max: nil, min: nil, man: nil, pt: nil)
        super(**{ id: id, max: max, min: min, man: man, pt: pt })
      end
      # rubocop:enable Naming/MethodParameterName

      # Returns whether the page break is a manual break.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def manual?
        [1, true].include?(man)
      end
    end

    # Cell comment or note (<comment>).
    #
    # +anchor+ places the comment box: the eight integers of the VML note's <x:Anchor>, that is
    # the left column, x offset, top row, y offset, right column, x offset, bottom row, and
    # y offset (0-based columns and rows, offsets in pixels). Without it, the box is placed
    # next to the cell.
    # @api public
    Comment = Data.define(:ref, :text, :author, :r_id, :shape_id, :author_id, :visible, :anchor) do
      include MetadataRecord

      private :visible

      #: (?ref: String?, ?text: String?, ?author: String?, ?r_id: String?, ?shape_id: Integer?, ?author_id: Integer?, ?visible: bool?, ?anchor: Array[Integer]?) -> void
      def initialize(ref: nil, text: nil, author: nil, r_id: nil, shape_id: nil, author_id: nil, visible: false, anchor: nil)
        raise ArgumentError, "comment anchor must be an Array of 8 Integers, got #{anchor.inspect}" unless anchor.nil? || (anchor.is_a?(Array) && anchor.size == 8 && anchor.all?(Integer))

        super(**{ ref: ref, text: text, author: author, r_id: r_id, shape_id: shape_id, author_id: author_id, visible: visible, anchor: anchor&.dup&.freeze })
      end

      # Returns whether the comment is visible by default.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def visible?
        visible ? true : false
      end

      # The anchor of a comment box without an explicit anchor: next to the cell at the
      # 0-based +column+ and +row+, two columns wide and four rows high.
      #
      # @param column [Integer]
      # @param row [Integer]
      # @return [Array<Integer>]
      # @api private
      #: (Integer column, Integer row) -> Array[Integer]
      def self.default_anchor(column, row)
        [column + 1, 15, row, 10, column + 3, 15, row + 4, 4]
      end

      # Returns whether the comment has no text.
      #
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty?
        text.nil? || text.empty?
      end

      # @return [String]
      # @api public
      #: () -> String
      def to_s
        text.to_s
      end
    end

    # Threaded comment (<threadedComment>).
    # @api public
    ThreadedComment = Data.define(:id, :ref, :text, :author, :author_id, :created_at, :parent_id, :resolved, :provider_id) do
      include MetadataRecord

      private :resolved

      #: (?id: String?, ?ref: String?, ?text: String?, ?author: String?, ?author_id: String?, ?created_at: untyped, ?parent_id: String?, ?resolved: bool?, ?provider_id: String?) -> void
      def initialize(
        id: nil, ref: nil, text: nil, author: nil,
        author_id: nil, created_at: nil,
        parent_id: nil, resolved: nil, provider_id: nil
      )
        super(**{ id: id, ref: ref, text: text, author: author, author_id: author_id, created_at: created_at, parent_id: parent_id, resolved: resolved, provider_id: provider_id })
      end

      # Returns whether the threaded comment is resolved.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def resolved?
        resolved ? true : false
      end

      # @return [String]
      # @api public
      #: () -> String
      def to_s
        text.to_s
      end
    end

    # Data validation rule (<dataValidation>).
    # @api public
    DataValidation = Data.define(
      :sqref, :type, :error_style, :operator, :allow_blank, :show_input_message,
      :show_error_message, :prompt_title, :prompt, :error_title, :error,
      :formula1, :formula2, :ime_mode, :show_drop_down
    ) do
      include MetadataRecord

      private :allow_blank, :show_input_message, :show_error_message, :show_drop_down

      #: (?sqref: String?, ?type: String?, ?error_style: String?, ?operator: String?, ?allow_blank: bool?, ?show_input_message: bool?, ?show_error_message: bool?, ?prompt_title: String?, ?prompt: String?, ?error_title: String?, ?error: String?, ?formula1: String?, ?formula2: String?, ?ime_mode: String?, ?show_drop_down: bool?) -> void
      def initialize(sqref: nil, type: nil, error_style: nil, operator: nil, allow_blank: nil, show_input_message: nil, show_error_message: nil, prompt_title: nil, prompt: nil, error_title: nil, error: nil, formula1: nil, formula2: nil, ime_mode: nil, show_drop_down: nil)
        super(**{
          sqref: sqref, type: type, error_style: error_style, operator: operator,
          allow_blank: allow_blank, show_input_message: show_input_message,
          show_error_message: show_error_message, prompt_title: prompt_title, prompt: prompt,
          error_title: error_title, error: error, formula1: formula1, formula2: formula2,
          ime_mode: ime_mode, show_drop_down: show_drop_down
        })
      end

      # Returns whether blank cells are allowed in validation.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def allow_blank?
        allow_blank ? true : false
      end

      # Returns whether the input message prompt is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_input_message?
        show_input_message ? true : false
      end

      # Returns whether the error alert message is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_error_message?
        show_error_message ? true : false
      end

      # Returns whether the in-cell dropdown is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_drop_down?
        show_drop_down ? true : false
      end
    end

    # Conditional formatting rule (<cfRule>).
    # @api public
    ConditionalFormatRule = Data.define(
      :type, :dxf_id, :priority, :operator, :formula, :text, :time_period,
      :rank, :std_dev, :above_average, :equal_average, :bottom, :percent,
      :stop_if_true, :color_scale, :data_bar, :icon_set, :options
    ) do
      include MetadataRecord

      private :above_average, :equal_average, :bottom, :percent, :stop_if_true

      #: (?type: (Symbol | String)?, ?dxf_id: Integer?, ?priority: Integer?, ?operator: (Symbol | String)?, ?formula: untyped, ?text: String?, ?time_period: (Symbol | String)?, ?rank: Integer?, ?std_dev: Integer?, ?above_average: bool?, ?equal_average: bool?, ?bottom: bool?, ?percent: bool?, ?stop_if_true: bool?, ?color_scale: Hash[Symbol, untyped]?, ?data_bar: Hash[Symbol, untyped]?, ?icon_set: Hash[Symbol, untyped]?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(type: nil, dxf_id: nil, priority: nil, operator: nil, formula: nil, text: nil, time_period: nil, rank: nil, std_dev: nil, above_average: nil, equal_average: nil, bottom: nil, percent: nil, stop_if_true: nil, color_scale: nil, data_bar: nil, icon_set: nil, options: nil, **extra)
        opts = (options || {}).merge(extra)
        super(**{
          type: type, dxf_id: dxf_id, priority: priority, operator: operator, formula: formula,
          text: text, time_period: time_period, rank: rank, std_dev: std_dev,
          above_average: above_average, equal_average: equal_average, bottom: bottom,
          percent: percent, stop_if_true: stop_if_true, color_scale: color_scale,
          data_bar: data_bar, icon_set: icon_set, options: opts.empty? ? nil : opts
        })
      end

      # Returns whether rule matches above average values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def above_average?
        above_average ? true : false
      end

      # Returns whether rule matches equal average values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def equal_average?
        equal_average ? true : false
      end

      # Returns whether rule matches bottom values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def bottom?
        bottom ? true : false
      end

      # Returns whether rank is evaluated as a percentage.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def percent?
        percent ? true : false
      end

      # Returns whether formatting evaluation stops if rule is true.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def stop_if_true?
        stop_if_true ? true : false
      end
    end

    # Conditional format definition (<conditionalFormatting> and <cfRule>).
    # @api public
    ConditionalFormat = Data.define(
      :sqref, :type, :dxf_id, :priority, :operator, :formula, :text, :time_period,
      :rank, :std_dev, :above_average, :equal_average, :bottom, :percent,
      :stop_if_true, :color_scale, :data_bar, :icon_set, :pivot, :rules,
      :options
    ) do
      include MetadataRecord

      private :above_average, :equal_average, :bottom, :percent, :stop_if_true, :pivot

      #: (?sqref: String?, ?type: (Symbol | String)?, ?dxf_id: Integer?, ?priority: Integer?, ?operator: (Symbol | String)?, ?formula: untyped, ?text: String?, ?time_period: (Symbol | String)?, ?rank: Integer?, ?std_dev: Integer?, ?above_average: bool?, ?equal_average: bool?, ?bottom: bool?, ?percent: bool?, ?stop_if_true: bool?, ?color_scale: Hash[Symbol, untyped]?, ?data_bar: Hash[Symbol, untyped]?, ?icon_set: Hash[Symbol, untyped]?, ?pivot: bool?, ?rules: Array[untyped]?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(sqref: nil, type: nil, dxf_id: nil, priority: nil, operator: nil, formula: nil, text: nil, time_period: nil, rank: nil, std_dev: nil, above_average: nil, equal_average: nil, bottom: nil, percent: nil, stop_if_true: nil, color_scale: nil, data_bar: nil, icon_set: nil, pivot: nil, rules: nil, options: nil, **extra)
        wrapped_rules = if rules && !rules.empty?
                          rules.map { |r| r.is_a?(ConditionalFormatRule) ? r : ConditionalFormatRule.new(**r) }
                        end
        opts = (options || {}).merge(extra)
        super(**{
          sqref: sqref, type: type, dxf_id: dxf_id, priority: priority, operator: operator, formula: formula,
          text: text, time_period: time_period, rank: rank, std_dev: std_dev,
          above_average: above_average, equal_average: equal_average, bottom: bottom,
          percent: percent, stop_if_true: stop_if_true, color_scale: color_scale,
          data_bar: data_bar, icon_set: icon_set, pivot: pivot, rules: wrapped_rules,
          options: opts.empty? ? nil : opts
        })
      end

      # Returns whether rule matches above average values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def above_average?
        above_average ? true : false
      end

      # Returns whether rule matches equal average values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def equal_average?
        equal_average ? true : false
      end

      # Returns whether rule matches bottom values.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def bottom?
        bottom ? true : false
      end

      # Returns whether rank is evaluated as a percentage.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def percent?
        percent ? true : false
      end

      # Returns whether formatting evaluation stops if rule is true.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def stop_if_true?
        stop_if_true ? true : false
      end

      # Returns whether conditional format applies to pivot table.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def pivot?
        pivot ? true : false
      end
    end

    # Column definition in an OpenXML table (<tableColumn>).
    # @api public
    TableColumn = Data.define(:id, :name, :totals_row_function, :totals_row_label, :totals_row_formula, :data_dxf_id, :calculated_column_formula, :data_cell_style) do
      include MetadataRecord

      #: (?id: Integer?, ?name: String?, ?totals_row_function: String?, ?totals_row_label: String?, ?totals_row_formula: String?, ?data_dxf_id: Integer?, ?calculated_column_formula: String?, ?data_cell_style: String?) -> void
      def initialize(id: nil, name: nil, totals_row_function: nil, totals_row_label: nil, totals_row_formula: nil, data_dxf_id: nil, calculated_column_formula: nil, data_cell_style: nil)
        super(**{
          id: id, name: name, totals_row_function: totals_row_function,
          totals_row_label: totals_row_label, totals_row_formula: totals_row_formula,
          data_dxf_id: data_dxf_id,
          calculated_column_formula: calculated_column_formula, data_cell_style: data_cell_style
        })
      end
    end

    # Table style information (<tableStyleInfo>).
    # @api public
    TableStyleInfo = Data.define(:name, :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes) do
      include MetadataRecord

      private :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes

      #: (?name: String?, ?show_first_column: bool?, ?show_last_column: bool?, ?show_row_stripes: bool?, ?show_column_stripes: bool?) -> void
      def initialize(name: nil, show_first_column: nil, show_last_column: nil, show_row_stripes: nil, show_column_stripes: nil)
        super(**{
          name: name, show_first_column: show_first_column, show_last_column: show_last_column,
          show_row_stripes: show_row_stripes, show_column_stripes: show_column_stripes
        })
      end

      # Returns whether first column is formatted distinctly.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_first_column?
        show_first_column ? true : false
      end

      # Returns whether last column is formatted distinctly.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_last_column?
        show_last_column ? true : false
      end

      # Returns whether row stripes are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_row_stripes?
        show_row_stripes ? true : false
      end

      # Returns whether column stripes are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_column_stripes?
        show_column_stripes ? true : false
      end
    end

    # OpenXML Table (<table id=... ref=...>).
    # @api public
    Table = Data.define(:id, :name, :display_name, :ref, :totals_row_count, :header_row_count, :columns, :style_info, :auto_filter) do
      include MetadataRecord

      #: (?id: Integer?, ?name: String?, ?display_name: String?, ?ref: String?, ?totals_row_count: Integer?, ?header_row_count: Integer?, ?columns: Array[Elements::TableColumn | Hash[Symbol, untyped]]?, ?style_info: (Elements::TableStyleInfo | Hash[Symbol, untyped])?, ?auto_filter: Hash[Symbol, untyped]?, ?totals_row_shown: bool?, ?header_row_shown: bool?, **untyped) -> void
      def initialize(id: nil, name: nil, display_name: nil, ref: nil, totals_row_count: nil, header_row_count: nil, columns: nil, style_info: nil, auto_filter: nil, totals_row_shown: nil, header_row_shown: nil, **extra)
        trc = totals_row_count
        if trc.nil?
          trs = totals_row_shown.nil? ? extra[:totals_row_shown] : totals_row_shown
          trc = trs ? 1 : 0 unless trs.nil?
        end
        hrc = header_row_count
        if hrc.nil?
          hrs = header_row_shown.nil? ? extra[:header_row_shown] : header_row_shown
          hrc = hrs ? 1 : 0 unless hrs.nil?
        end
        wrapped_cols = (columns || []).map { |c| c.is_a?(TableColumn) ? c : TableColumn.new(**c) }
        wrapped_si = style_info.is_a?(Hash) ? TableStyleInfo.new(**style_info) : style_info
        super(**{
          id: id, name: name, display_name: display_name, ref: ref,
          totals_row_count: trc, header_row_count: hrc,
          columns: wrapped_cols, style_info: wrapped_si, auto_filter: auto_filter
        })
      end

      # Builds the style_info of a table hash whose style is given as :style: a style name with
      # the show_* flags at the root of the hash, or a Hash or TableStyleInfo as the reader returns.
      #
      # @param table [Hash{Symbol => Object}]
      # @return [Hash, Elements::TableStyleInfo, nil]
      # @api private
      #: (Hash[Symbol, untyped] table) -> (Hash[Symbol, untyped] | Elements::TableStyleInfo)?
      def self.style_info_from(table)
        style = table[:style]
        return style if style.nil? || style.is_a?(Hash) || style.is_a?(TableStyleInfo)

        { name: style.to_s }.merge(table.slice(:show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes))
      end

      # Returns whether the table totals row is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def totals_row_shown?
        (totals_row_count || 0).positive?
      end

      # Returns whether the table header row is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def header_row_shown?
        header_row_count.nil? || header_row_count.positive?
      end

      # Returns the table style name.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def style_name
        style_info&.name
      end

      # Returns whether first column is formatted distinctly.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_first_column?
        style_info ? style_info.show_first_column? : false
      end

      # Returns whether last column is formatted distinctly.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_last_column?
        style_info ? style_info.show_last_column? : false
      end

      # Returns whether row stripes are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_row_stripes?
        style_info ? style_info.show_row_stripes? : false
      end

      # Returns whether column stripes are shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_column_stripes?
        style_info ? style_info.show_column_stripes? : false
      end
    end

    # Single sparkline definition (<x14:sparkline>).
    # @api public
    Sparkline = Data.define(:sqref, :formula) do
      include MetadataRecord

      #: (?sqref: String?, ?formula: String?) -> void
      def initialize(sqref: nil, formula: nil)
        super(**{ sqref: sqref, formula: formula })
      end
    end

    # Sparkline group (<x14:sparklineGroup>).
    # @api public
    SparklineGroup = Data.define(
      :type, :color_series, :color_negative, :color_axis, :color_markers,
      :color_first, :color_last, :color_high, :color_low, :sparklines,
      :display_empty_cells_as, :markers, :high, :low, :first, :last, :negative,
      :manual_max, :manual_min, :line_weight
    ) do
      include MetadataRecord

      private :markers, :high, :low, :first, :last, :negative

      #: (?type: String?, ?color_series: String?, ?color_negative: String?, ?color_axis: String?, ?color_markers: String?, ?color_first: String?, ?color_last: String?, ?color_high: String?, ?color_low: String?, ?sparklines: Array[Elements::Sparkline | Hash[Symbol, untyped]]?, ?display_empty_cells_as: String?, ?markers: bool?, ?high: bool?, ?low: bool?, ?first: bool?, ?last: bool?, ?negative: bool?, ?manual_max: (Integer | Float)?, ?manual_min: (Integer | Float)?, ?line_weight: (Integer | Float)?, **untyped) -> void
      def initialize(type: nil, color_series: nil, color_negative: nil, color_axis: nil, color_markers: nil, color_first: nil, color_last: nil, color_high: nil, color_low: nil, sparklines: nil, display_empty_cells_as: nil, markers: nil, high: nil, low: nil, first: nil, last: nil, negative: nil, manual_max: nil, manual_min: nil, line_weight: nil, **extra)
        wrapped_sps = (sparklines || []).map { |s| s.is_a?(Sparkline) ? s : Sparkline.new(**s) }
        lw = line_weight || extra[:line_weight]
        super(**{
          type: type, color_series: color_series, color_negative: color_negative,
          color_axis: color_axis, color_markers: color_markers, color_first: color_first,
          color_last: color_last, color_high: color_high, color_low: color_low,
          sparklines: wrapped_sps, display_empty_cells_as: display_empty_cells_as,
          markers: markers, high: high, low: low, first: first, last: last, negative: negative,
          manual_max: manual_max, manual_min: manual_min, line_weight: lw
        })
      end

      # Returns whether data markers are displayed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def markers?
        markers ? true : false
      end

      # Returns whether high point markers are highlighted.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def high?
        high ? true : false
      end

      # Returns whether low point markers are highlighted.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def low?
        low ? true : false
      end

      # Returns whether first point markers are highlighted.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def first?
        first ? true : false
      end

      # Returns whether last point markers are highlighted.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def last?
        last ? true : false
      end

      # Returns whether negative point markers are highlighted.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def negative?
        negative ? true : false
      end
    end

    # Protected range configuration (<protectedRange>).
    # @api public
    ProtectedRange = Data.define(:name, :sqref, :security_descriptor, :password, :algorithm_name, :hash_value, :salt_value, :spin_count) do
      include MetadataRecord

      #: (?name: String?, ?sqref: String?, ?security_descriptor: String?, ?password: String?, ?algorithm_name: String?, ?hash_value: String?, ?salt_value: String?, ?spin_count: Integer?) -> void
      def initialize(name: nil, sqref: nil, security_descriptor: nil, password: nil, algorithm_name: nil, hash_value: nil, salt_value: nil, spin_count: nil)
        super(**{
          name: name, sqref: sqref, security_descriptor: security_descriptor, password: password,
          algorithm_name: algorithm_name, hash_value: hash_value, salt_value: salt_value, spin_count: spin_count
        })
      end
    end

    # Ignored error rule (<ignoredError>).
    # @api public
    IgnoredError = Data.define(
      :sqref, :eval_error, :two_digit_text_year, :number_stored_as_text,
      :formula, :formula_range, :unlocked_formula, :empty_cell_reference,
      :list_data_validation, :calculated_column
    ) do
      include MetadataRecord

      private :eval_error, :two_digit_text_year, :number_stored_as_text, :formula, :formula_range, :unlocked_formula, :empty_cell_reference, :list_data_validation, :calculated_column

      #: (?sqref: String?, ?eval_error: bool?, ?two_digit_text_year: bool?, ?number_stored_as_text: bool?, ?formula: bool?, ?formula_range: bool?, ?unlocked_formula: bool?, ?empty_cell_reference: bool?, ?list_data_validation: bool?, ?calculated_column: bool?) -> void
      def initialize(sqref: nil, eval_error: nil, two_digit_text_year: nil, number_stored_as_text: nil, formula: nil, formula_range: nil, unlocked_formula: nil, empty_cell_reference: nil, list_data_validation: nil, calculated_column: nil)
        super(**{
          sqref: sqref, eval_error: eval_error, two_digit_text_year: two_digit_text_year,
          number_stored_as_text: number_stored_as_text, formula: formula, formula_range: formula_range,
          unlocked_formula: unlocked_formula, empty_cell_reference: empty_cell_reference,
          list_data_validation: list_data_validation, calculated_column: calculated_column
        })
      end

      # Returns whether evaluation error check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def eval_error?
        eval_error ? true : false
      end

      # Returns whether two-digit text year check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def two_digit_text_year?
        two_digit_text_year ? true : false
      end

      # Returns whether number stored as text check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def number_stored_as_text?
        number_stored_as_text ? true : false
      end

      # Returns whether formula syntax check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def formula?
        formula ? true : false
      end

      # Returns whether formula range check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def formula_range?
        formula_range ? true : false
      end

      # Returns whether unlocked formula check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def unlocked_formula?
        unlocked_formula ? true : false
      end

      # Returns whether empty cell reference check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def empty_cell_reference?
        empty_cell_reference ? true : false
      end

      # Returns whether list data validation error check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def list_data_validation?
        list_data_validation ? true : false
      end

      # Returns whether calculated column error check is ignored.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def calculated_column?
        calculated_column ? true : false
      end
    end

    # Drawing chart entry (<c:chart>).
    # @api public
    Chart = Data.define(
      :r_id, :name, :title, :type, :anchor, :from, :to, :series, :legend,
      :category_axis, :value_axis, :secondary_value_axis, :grouping,
      :gap_width, :overlap, :target, :from_column, :from_row, :to_column,
      :to_row, :data_labels, :options
    ) do
      include MetadataRecord

      #: (?r_id: String?, ?name: String?, ?title: String?, ?type: String?, ?anchor: String?, ?from: untyped, ?to: untyped, ?series: Array[untyped]?, ?legend: untyped, ?category_axis: untyped, ?value_axis: untyped, ?secondary_value_axis: untyped, ?grouping: untyped, ?gap_width: untyped, ?overlap: untyped, ?target: untyped, ?from_column: untyped, ?from_row: untyped, ?to_column: untyped, ?to_row: untyped, ?data_labels: untyped?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(r_id: nil, name: nil, title: nil, type: nil, anchor: nil, from: nil, to: nil, series: nil, legend: nil, category_axis: nil, value_axis: nil, secondary_value_axis: nil, grouping: nil, gap_width: nil, overlap: nil, target: nil, from_column: nil, from_row: nil, to_column: nil, to_row: nil, data_labels: nil, options: nil, **extra)
        opts = (options || {}).merge(extra)
        super(**{
          r_id: r_id, name: name, title: title, type: type, anchor: anchor, from: from, to: to,
          series: series, legend: legend, category_axis: category_axis, value_axis: value_axis,
          secondary_value_axis: secondary_value_axis, grouping: grouping, gap_width: gap_width,
          overlap: overlap, target: target, from_column: from_column, from_row: from_row,
          to_column: to_column, to_row: to_row, data_labels: data_labels,
          options: opts.empty? ? nil : opts
        })
      end
    end

    # Drawing shape entry (<xdr:sp>).
    # @api public
    Shape = Data.define(:id, :name, :type, :anchor, :from, :to, :text, :preset, :width, :height, :left, :top, :options) do
      include MetadataRecord

      #: (?id: Integer?, ?name: String?, ?type: String?, ?anchor: String?, ?from: untyped, ?to: untyped, ?text: String?, ?preset: String?, ?width: (Float | Integer)?, ?height: (Float | Integer)?, ?left: (Float | Integer)?, ?top: (Float | Integer)?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(id: nil, name: nil, type: nil, anchor: nil, from: nil, to: nil, text: nil, preset: nil, width: nil, height: nil, left: nil, top: nil, options: nil, **extra)
        opts = (options || {}).merge(extra)
        super(**{
          id: id, name: name, type: type, anchor: anchor, from: from, to: to,
          text: text, preset: preset, width: width, height: height, left: left, top: top,
          options: opts.empty? ? nil : opts
        })
      end
    end

    # Pivot table entry (<pivotTableDefinition>).
    # @api public
    PivotTable = Data.define(
      :name, :ref, :source_ref, :row_fields, :column_fields,
      :data_fields, :page_fields, :data_caption, :page_wrap,
      :column_grand_totals, :row_grand_totals, :show_row_headers,
      :show_column_headers, :show_row_stripes, :show_column_stripes,
      :cache, :options
    ) do
      include MetadataRecord

      private :column_grand_totals, :row_grand_totals, :show_row_headers, :show_column_headers, :show_row_stripes, :show_column_stripes

      #: (?name: String?, ?ref: String?, ?source_ref: String?, ?row_fields: Array[untyped]?, ?column_fields: Array[untyped]?, ?data_fields: Array[untyped]?, ?page_fields: Array[untyped]?, ?data_caption: String?, ?page_wrap: Integer?, ?column_grand_totals: bool?, ?row_grand_totals: bool?, ?show_row_headers: bool?, ?show_column_headers: bool?, ?show_row_stripes: bool?, ?show_column_stripes: bool?, ?cache: Hash[Symbol, untyped]?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(
        name: nil, ref: nil, source_ref: nil,
        row_fields: nil, column_fields: nil, data_fields: nil, page_fields: nil,
        data_caption: nil, page_wrap: nil, column_grand_totals: nil, row_grand_totals: nil,
        show_row_headers: nil, show_column_headers: nil, show_row_stripes: nil, show_column_stripes: nil,
        cache: nil, options: nil, **extra
      )
        cols = column_fields || []
        merged_opts = (options || {}).merge(extra).freeze
        super(**{
          name: name, ref: ref, source_ref: source_ref,
          row_fields: row_fields, column_fields: cols, data_fields: data_fields, page_fields: page_fields,
          data_caption: data_caption, page_wrap: page_wrap,
          column_grand_totals: column_grand_totals, row_grand_totals: row_grand_totals,
          show_row_headers: show_row_headers, show_column_headers: show_column_headers,
          show_row_stripes: show_row_stripes, show_column_stripes: show_column_stripes,
          cache: cache, options: merged_opts.empty? ? nil : merged_opts
        })
      end

      # Returns whether column grand totals are displayed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def column_grand_totals?
        column_grand_totals ? true : false
      end

      # Returns whether row grand totals are displayed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def row_grand_totals?
        row_grand_totals ? true : false
      end

      # Returns whether row headers formatting is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_row_headers?
        show_row_headers ? true : false
      end

      # Returns whether column headers formatting is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_column_headers?
        show_column_headers ? true : false
      end

      # Returns whether row stripe banding is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_row_stripes?
        show_row_stripes ? true : false
      end

      # Returns whether column stripe banding is shown.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_column_stripes?
        show_column_stripes ? true : false
      end
    end

    # Header or footer embedded image.
    # @api public
    HeaderFooterImage = Data.define(:position, :r_id, :path, :extension, :data, :content_type) do
      include MetadataRecord

      #: (?position: String?, ?r_id: String?, ?path: String?, ?extension: String?, ?data: String?, ?content_type: String?) -> void
      def initialize(position: nil, r_id: nil, path: nil, extension: nil, data: nil, content_type: nil)
        super(**{ position: position, r_id: r_id, path: path, extension: extension, data: data, content_type: content_type })
      end
    end

    # Font style definition in stylesheet (<font>).
    # @api public
    # rubocop:disable Lint/DataDefineOverride
    Font = Data.define(:name, :size, :bold, :italic, :underline, :strike, :color, :family, :scheme, :outline, :shadow, :condense, :extend, :vert_align) do
      # rubocop:enable Lint/DataDefineOverride
      include MetadataRecord

      private :bold, :italic, :strike, :outline, :shadow, :condense, :extend

      #: (?name: String?, ?size: (Integer | Float)?, ?bold: bool?, ?italic: bool?, ?underline: (String | bool)?, ?strike: bool?, ?color: Hash[Symbol, untyped]?, ?family: Integer?, ?scheme: String?, ?outline: bool?, ?shadow: bool?, ?condense: bool?, ?extend: bool?, ?vert_align: String?) -> void
      def initialize(name: nil, size: nil, bold: nil, italic: nil, underline: nil, strike: nil, color: nil, family: nil, scheme: nil, outline: nil, shadow: nil, condense: nil, extend: nil, vert_align: nil)
        super(**{ name: name, size: size, bold: bold, italic: italic, underline: underline, strike: strike, color: color, family: family, scheme: scheme, outline: outline, shadow: shadow, condense: condense, extend: extend, vert_align: vert_align })
      end

      # Returns whether the font is bold.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def bold?
        bold ? true : false
      end

      # Returns whether the font is italic.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def italic?
        italic ? true : false
      end

      # Returns whether the font has strikethrough.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def strike?
        strike ? true : false
      end

      # Returns whether the font has outline effect.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def outline?
        outline ? true : false
      end

      # Returns whether the font has shadow effect.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def shadow?
        shadow ? true : false
      end

      # Returns whether the font is condensed.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def condense?
        condense ? true : false
      end

      # Returns whether the font is extended.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def extend?
        extend ? true : false
      end
    end

    # Fill style definition in stylesheet (<fill>).
    #
    # +gradient+ describes a <gradientFill> in place of a pattern: { type:, degree:, left:, right:,
    # top:, bottom:, stops: [{ position:, color: { rgb: | theme:, tint: } }] }.
    # @api public
    Fill = Data.define(:pattern_type, :fg_color, :bg_color, :gradient) do
      include MetadataRecord

      #: (?pattern_type: String?, ?fg_color: Hash[Symbol, untyped]?, ?bg_color: Hash[Symbol, untyped]?, ?gradient: Hash[Symbol, untyped]?) -> void
      def initialize(pattern_type: nil, fg_color: nil, bg_color: nil, gradient: nil)
        super(**{ pattern_type: pattern_type, fg_color: fg_color, bg_color: bg_color, gradient: gradient })
      end
    end

    # Border style definition in stylesheet (<border>).
    # @api public
    Border = Data.define(:left, :right, :top, :bottom, :diagonal, :diagonal_up, :diagonal_down) do
      include MetadataRecord

      private :diagonal_up, :diagonal_down

      #: (?left: Hash[Symbol, untyped]?, ?right: Hash[Symbol, untyped]?, ?top: Hash[Symbol, untyped]?, ?bottom: Hash[Symbol, untyped]?, ?diagonal: Hash[Symbol, untyped]?, ?diagonal_up: bool?, ?diagonal_down: bool?) -> void
      def initialize(left: nil, right: nil, top: nil, bottom: nil, diagonal: nil, diagonal_up: nil, diagonal_down: nil)
        super(**{ left: left, right: right, top: top, bottom: bottom, diagonal: diagonal, diagonal_up: diagonal_up, diagonal_down: diagonal_down })
      end

      # Returns whether diagonal up border is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def diagonal_up?
        diagonal_up ? true : false
      end

      # Returns whether diagonal down border is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def diagonal_down?
        diagonal_down ? true : false
      end
    end

    # Alignment style definition in stylesheet (<alignment>).
    # @api public
    Alignment = Data.define(
      :horizontal, :vertical, :wrap_text, :text_rotation, :indent,
      :relative_indent, :shrink_to_fit, :reading_order, :justify_last_line
    ) do
      include MetadataRecord

      private :wrap_text, :shrink_to_fit, :justify_last_line

      #: (?horizontal: (String | Symbol)?, ?vertical: (String | Symbol)?, ?wrap_text: bool?, ?text_rotation: Numeric?, ?indent: Integer?, ?relative_indent: Integer?, ?shrink_to_fit: bool?, ?reading_order: Integer?, ?justify_last_line: bool?) -> void
      def initialize(horizontal: nil, vertical: nil, wrap_text: nil, text_rotation: nil, indent: nil, relative_indent: nil, shrink_to_fit: nil, reading_order: nil, justify_last_line: nil)
        super(**{ horizontal: horizontal, vertical: vertical, wrap_text: wrap_text, text_rotation: text_rotation, indent: indent, relative_indent: relative_indent, shrink_to_fit: shrink_to_fit, reading_order: reading_order, justify_last_line: justify_last_line })
      end

      # Returns whether text wrapping is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def wrap_text?
        wrap_text ? true : false
      end

      # Returns whether shrink to fit is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def shrink_to_fit?
        shrink_to_fit ? true : false
      end

      # Returns whether justify last line is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def justify_last_line?
        justify_last_line ? true : false
      end
    end

    # Protection style definition in stylesheet (<protection>).
    # @api public
    Protection = Data.define(:locked, :hidden) do
      include MetadataRecord

      private :locked, :hidden

      #: (?locked: bool?, ?hidden: bool?) -> void
      def initialize(locked: nil, hidden: nil)
        super(**{ locked: locked, hidden: hidden })
      end

      # Returns whether the cell is locked.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def locked?
        locked ? true : false
      end

      # Returns whether the cell formula is hidden.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def hidden?
        hidden ? true : false
      end
    end

    # Cell format record in stylesheet (<xf> in <cellXfs> or <cellStyleXfs>).
    # @api public
    CellXf = Data.define(
      :num_fmt_id, :font_id, :fill_id, :border_id, :xf_id,
      :apply_number_format, :apply_font, :apply_fill, :apply_border,
      :apply_alignment, :apply_protection, :alignment, :protection,
      :quote_prefix, :pivot_button
    ) do
      include MetadataRecord

      private :apply_number_format, :apply_font, :apply_fill, :apply_border, :apply_alignment, :apply_protection, :quote_prefix, :pivot_button

      #: (?num_fmt_id: Integer?, ?font_id: Integer?, ?fill_id: Integer?, ?border_id: Integer?, ?xf_id: Integer?, ?apply_number_format: bool?, ?apply_font: bool?, ?apply_fill: bool?, ?apply_border: bool?, ?apply_alignment: bool?, ?apply_protection: bool?, ?alignment: (Elements::Alignment | Hash[Symbol, untyped])?, ?protection: (Elements::Protection | Hash[Symbol, untyped])?, ?quote_prefix: bool?, ?pivot_button: bool?) -> void
      def initialize(
        num_fmt_id: nil, font_id: nil, fill_id: nil, border_id: nil, xf_id: nil,
        apply_number_format: nil, apply_font: nil, apply_fill: nil, apply_border: nil,
        apply_alignment: nil, apply_protection: nil, alignment: nil, protection: nil,
        quote_prefix: nil, pivot_button: nil
      )
        align_obj = alignment.is_a?(Hash) ? Alignment.new(**alignment) : alignment
        prot_obj = protection.is_a?(Hash) ? Protection.new(**protection) : protection
        super(**{
          num_fmt_id: num_fmt_id, font_id: font_id, fill_id: fill_id, border_id: border_id,
          xf_id: xf_id, apply_number_format: apply_number_format, apply_font: apply_font,
          apply_fill: apply_fill, apply_border: apply_border, apply_alignment: apply_alignment,
          apply_protection: apply_protection, alignment: align_obj, protection: prot_obj,
          quote_prefix: quote_prefix, pivot_button: pivot_button
        })
      end

      # Returns whether quote prefix is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def quote_prefix?
        quote_prefix ? true : false
      end

      # Returns whether pivot button is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def pivot_button?
        pivot_button ? true : false
      end

      # Returns whether number format is applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_number_format?
        apply_number_format ? true : false
      end

      # Returns whether font is applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_font?
        apply_font ? true : false
      end

      # Returns whether fill is applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_fill?
        apply_fill ? true : false
      end

      # Returns whether border is applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_border?
        apply_border ? true : false
      end

      # Returns whether alignment is applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_alignment?
        apply_alignment ? true : false
      end

      # Returns whether protection is applied.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def apply_protection?
        apply_protection ? true : false
      end
    end

    # Differential format in stylesheet (<dxf> in <dxfs>).
    # @api public
    Dxf = Data.define(:font, :fill, :border, :number_format, :alignment, :protection) do
      include MetadataRecord

      #: (?font: (Elements::Font | Hash[Symbol, untyped])?, ?fill: (Elements::Fill | Hash[Symbol, untyped])?, ?border: (Elements::Border | Hash[Symbol, untyped])?, ?number_format: Hash[Symbol, untyped]?, ?alignment: (Elements::Alignment | Hash[Symbol, untyped])?, ?protection: (Elements::Protection | Hash[Symbol, untyped])?) -> void
      def initialize(font: nil, fill: nil, border: nil, number_format: nil, alignment: nil, protection: nil)
        f = font.is_a?(Hash) ? Font.new(**font) : font
        # Differential fills from the styles reader name the pattern :pattern only.
        fill = fill.merge(pattern_type: fill[:pattern]).except(:pattern) if fill.is_a?(Hash) && fill.key?(:pattern) && !fill.key?(:pattern_type)
        fi = fill.is_a?(Hash) ? Fill.new(**fill) : fill
        b = border.is_a?(Hash) ? Border.new(**border) : border
        align = alignment.is_a?(Hash) ? Alignment.new(**alignment) : alignment
        prot = protection.is_a?(Hash) ? Protection.new(**protection) : protection
        super(**{ font: f, fill: fi, border: b, number_format: number_format, alignment: align, protection: prot })
      end
    end

    # Named cell style in stylesheet (<cellStyle> in <cellStyles>).
    # @api public
    CellStyle = Data.define(:name, :xf_id, :builtin_id) do
      include MetadataRecord

      #: (?name: String?, ?xf_id: Integer?, ?builtin_id: Integer?) -> void
      def initialize(name: nil, xf_id: nil, builtin_id: nil)
        super(**{ name: name, xf_id: xf_id, builtin_id: builtin_id })
      end
    end

    # Color definition in stylesheet (<color>).
    # @api public
    Color = Data.define(:rgb, :indexed, :theme, :tint) do
      include MetadataRecord

      #: (?rgb: String?, ?indexed: Integer?, ?theme: Integer?, ?tint: Float?) -> void
      def initialize(rgb: nil, indexed: nil, theme: nil, tint: nil)
        super(**{ rgb: rgb, indexed: indexed, theme: theme, tint: tint })
      end
    end

    # Old key names that readers and builders still emit next to the member they were renamed to
    # (see StrictKeywords).
    FreezePane::KEY_ALIASES = { col: :column }.freeze
    SplitPane::KEY_ALIASES = { row: :y_split, column: :x_split, col: :x_split }.freeze
    SheetFormat::KEY_ALIASES = { default_col_width: :default_column_width, base_col_width: :base_column_width, outline_level_col: :outline_level_column }.freeze
    Sparkline::KEY_ALIASES = { data_ref: :formula, location_ref: :sqref }.freeze
    AppProperties::KEY_ALIASES = { application: :app_name }.freeze
    WorkbookProperties::KEY_ALIASES = { active_sheet_index: :active_sheet }.freeze
    TableColumn::KEY_ALIASES = { data_formula: :calculated_column_formula }.freeze
    Font::KEY_ALIASES = { sz: :size }.freeze
    Fill::KEY_ALIASES = { pattern: :pattern_type }.freeze
    BackgroundImage::KEY_ALIASES = { ext: :extension, filename: :path, file_data: :data }.freeze
    WorkbookProtection::KEY_ALIASES = { password: :workbook_password }.freeze
    ThreadedComment::KEY_ALIASES = { dt: :created_at, timestamp: :created_at, done: :resolved, person_id: :author_id, user_id: :author_id }.freeze
  end
end
