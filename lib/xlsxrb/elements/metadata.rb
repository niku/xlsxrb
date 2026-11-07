# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Elements
    # Shared mixin for metadata Data value objects providing Hash compatibility.
    #
    # @api public
    module MetadataRecord
      # Accesses an attribute by Symbol or String key, supporting key aliases.
      #
      # @deprecated Hash-like access via {#[]} is deprecated and will be removed in 0.2.0; use method calls instead.
      # @param key [Symbol, String]
      # @return [Object, nil]
      # @api public
      #: (Symbol | String key) -> untyped
      def [](key)
        warn "[DEPRECATION] Hash access on #{self.class} (`#[]`) is deprecated and will be removed in 0.2.0; use method calls instead.", uplevel: 1
        _lookup_key(key)
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
            val.map { |v| v.respond_to?(:to_h) && !v.is_a?(String) ? v.to_h : v }
          elsif val.respond_to?(:to_h) && !val.is_a?(String) && !val.is_a?(Hash)
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

      # Checks whether a key is present with a non-nil value.
      #
      # @deprecated {#key?} is deprecated and will be removed in 0.2.0; check attribute presence directly.
      # @param key [Symbol, String]
      # @return [Boolean]
      # @api public
      #: (Symbol | String key) -> bool
      def key?(key)
        warn "[DEPRECATION] `#{self.class}#key?` is deprecated and will be removed in 0.2.0; check attribute presence directly.", uplevel: 1
        _has_key?(key)
      end
      alias has_key? key?

      # Digs into nested metadata by key path.
      #
      # @deprecated {#dig} is deprecated and will be removed in 0.2.0; use method chaining instead.
      # @param head [Symbol, String]
      # @param rest [Array<Symbol, String>]
      # @return [Object, nil]
      # @api public
      #: (Symbol | String head, *(Symbol | String) rest) -> untyped
      def dig(head, *rest)
        warn "[DEPRECATION] `#{self.class}#dig` is deprecated and will be removed in 0.2.0; use method chaining instead.", uplevel: 1
        val = _lookup_key(head)
        return nil if val.nil?
        return val if rest.empty?

        val.respond_to?(:dig) ? val.dig(*rest) : nil
      end

      # Fetches a key with an optional default.
      #
      # @deprecated {#fetch} is deprecated and will be removed in 0.2.0; use method calls instead.
      # @param key [Symbol, String]
      # @param args [Array<Object>]
      # @return [Object, nil]
      # @api public
      #: (Symbol | String key, *untyped args) ?{ (Symbol | String) -> untyped } -> untyped
      def fetch(key, *args)
        warn "[DEPRECATION] `#{self.class}#fetch` is deprecated and will be removed in 0.2.0; use method calls instead.", uplevel: 1
        val = _lookup_key(key)
        if !val.nil? || (_has_key?(key) && members.include?(key.to_sym))
          val
        elsif block_given?
          yield(key)
        elsif !args.empty?
          args.first
        else
          raise KeyError, "key not found: #{key.inspect}"
        end
      end

      # Merges non-nil attributes with another Hash or MetadataRecord.
      #
      # @deprecated {#merge} is deprecated and will be removed in 0.2.0; use {#with} or {#to_h} instead.
      # @param other [Hash, MetadataRecord]
      # @return [Hash{Symbol => Object}]
      # @api public
      #: (untyped other) -> Hash[Symbol, untyped]
      def merge(other)
        warn "[DEPRECATION] `#{self.class}#merge` is deprecated and will be removed in 0.2.0; use `#with` or `#to_h.merge` instead.", uplevel: 1
        h = to_h
        other_h = other.respond_to?(:to_h) ? other.to_h : other
        h.merge(other_h)
      end

      alias to_hash to_h

      # Checks equality against another instance, a Hash, or a scalar.
      #
      # @param other [Object]
      # @return [Boolean]
      # @api public
      #: (untyped other) -> bool
      def ==(other)
        if other.is_a?(Hash)
          warn "[DEPRECATION] Comparing `#{self.class}` with a Hash is deprecated and will be removed in 0.2.0; compare attributes directly or use `#to_h`.", uplevel: 1
          to_h == other.transform_keys(&:to_sym).compact
        else
          super
        end
      end

      private

      def _lookup_key(key)
        k = _normalize_key(key)
        return public_send(k) if members.include?(k)

        if members.include?(:options)
          opts = options
          return opts[k] || opts[key] if opts.is_a?(Hash)
        end
        nil
      end

      def _has_key?(key)
        k = _normalize_key(key)
        return true if members.include?(k) && !public_send(k).nil?

        if members.include?(:options)
          opts = options
          return true if opts.is_a?(Hash) && (opts.key?(k) || opts.key?(key))
        end
        false
      end

      def _normalize_key(key)
        k = key.is_a?(Symbol) ? key : key.to_s.to_sym
        k = :size if k == :sz
        k = :number_format if k == :num_fmt
        k = :page_setup_properties if k == :page_setup_pr
        k = :pattern_type if k == :pattern
        k = :column if k == :col
        k = :columns if k == :cols
        k = :ref if k == :cell
        k = :sheet_calc_properties if %i[sheet_calc_pr calc_pr].include?(k)
        k = :sqref if %i[location_ref location].include?(k)
        k = :formula if %i[formulas data_ref f].include?(k)
        k = :dxf_id if k == :format_id
        k = :default_column_width if k == :default_col_width
        k = :outline_level_column if k == :outline_level_col
        k = :base_column_width if k == :base_col_width
        k = :from_column if k == :from_col
        k = :to_column if k == :to_col
        k = :column_id if k == :col_id
        k = :active_sheet if k == :active_tab
        k = :first_sheet if k == :first_tab
        k = :background_image if %i[background_picture picture].include?(k)
        k = :local_sheet_name if k == :sheet
        k = :local_sheet_id if k == :sheet_id
        k = :value if k == :range
        k = :app_name if k == :application
        k
      end
    end

    # Core document properties (docProps/core.xml).
    # @api public
    CoreProperties = Data.define(
      :title, :subject, :creator, :keywords, :description, :last_modified_by,
      :revision, :created, :modified, :category, :content_status
    ) do
      include MetadataRecord

      #: (?title: String?, ?subject: String?, ?creator: String?, ?keywords: String?, ?description: String?, ?last_modified_by: String?, ?revision: String?, ?created: untyped, ?modified: untyped, ?category: String?, ?content_status: String?, **untyped) -> void
      def initialize(title: nil, subject: nil, creator: nil, keywords: nil, description: nil, last_modified_by: nil, revision: nil, created: nil, modified: nil, category: nil, content_status: nil, **_extra)
        super(
          title: title, subject: subject, creator: creator, keywords: keywords, description: description,
          last_modified_by: last_modified_by, revision: revision, created: created, modified: modified,
          category: category, content_status: content_status
        )
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

      #: (?app_name: String?, ?app_version: String?, ?doc_security: Integer?, ?scale_crop: bool?, ?heading_pairs: Array[untyped]?, ?titles_of_parts: Array[untyped]?, ?manager: String?, ?company: String?, ?links_up_to_date: bool?, ?shared_doc: bool?, ?hyperlink_base: String?, ?hyperlinks_changed: bool?, **untyped) -> void
      def initialize(app_name: nil, app_version: nil, doc_security: nil, scale_crop: nil, heading_pairs: nil, titles_of_parts: nil, manager: nil, company: nil, links_up_to_date: nil, shared_doc: nil, hyperlink_base: nil, hyperlinks_changed: nil, **extra)
        app_name ||= extra[:application]
        super(
          app_name: app_name, app_version: app_version, doc_security: doc_security, scale_crop: scale_crop,
          heading_pairs: heading_pairs, titles_of_parts: titles_of_parts, manager: manager, company: company,
          links_up_to_date: links_up_to_date, shared_doc: shared_doc, hyperlink_base: hyperlink_base,
          hyperlinks_changed: hyperlinks_changed
        )
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
      :date1904, :code_name, :active_sheet, :first_sheet, :first_tab,
      :show_sheet_tabs, :tab_ratio, :full_calc_on_load, :filter_privacy,
      :default_theme_version, :refresh_all
    ) do
      include MetadataRecord

      #: (?date1904: bool?, ?code_name: String?, ?active_sheet: Integer?, ?active_tab: Integer?, ?first_sheet: Integer?, ?first_tab: Integer?, ?show_sheet_tabs: bool?, ?tab_ratio: Integer?, ?full_calc_on_load: bool?, ?filter_privacy: bool?, ?default_theme_version: String?, ?refresh_all: bool?, **untyped) -> void
      def initialize(date1904: nil, code_name: nil, active_sheet: nil, active_tab: nil, first_sheet: nil, first_tab: nil, show_sheet_tabs: nil, tab_ratio: nil, full_calc_on_load: nil, filter_privacy: nil, default_theme_version: nil, refresh_all: nil, **extra)
        if (!active_tab.nil? || extra.key?(:active_tab)) && active_sheet.nil?
          warn "[Xlsxrb DEPRECATION] WorkbookProperties.new(active_tab:) is deprecated and will be removed in v0.2.0. Use active_sheet: instead.", uplevel: 1
          active_sheet = active_tab || extra[:active_tab]
        end
        ft = first_tab || first_sheet
        super(
          date1904: date1904, code_name: code_name, active_sheet: active_sheet, first_sheet: ft, first_tab: ft,
          show_sheet_tabs: show_sheet_tabs, tab_ratio: tab_ratio, full_calc_on_load: full_calc_on_load,
          filter_privacy: filter_privacy, default_theme_version: default_theme_version, refresh_all: refresh_all
        )
      end

      undef_method :first_tab

      # @deprecated Use {#first_sheet} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def first_tab
        warn "[Xlsxrb DEPRECATION] WorkbookProperties#first_tab is deprecated; use #first_sheet instead.", uplevel: 1
        first_sheet
      end

      # @deprecated Use {#active_sheet} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def active_tab
        warn "[Xlsxrb DEPRECATION] WorkbookProperties#active_tab is deprecated and will be removed in v0.2.0. Use #active_sheet instead.", uplevel: 1
        active_sheet
      end

      # Returns whether the 1904 date system is used.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def date1904?
        date1904 ? true : false
      end

      # Returns whether sheet tabs are shown in the workbook window.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_sheet_tabs?
        show_sheet_tabs != false
      end

      # Returns whether full calculation is performed when loading the workbook.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def full_calc_on_load?
        full_calc_on_load ? true : false
      end

      # Returns whether external data connections are refreshed when opening.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def refresh_all?
        refresh_all ? true : false
      end

      # Returns whether filter privacy is enabled.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def filter_privacy?
        filter_privacy ? true : false
      end
    end

    # Workbook protection settings (<workbookProtection>).
    # @api public
    WorkbookProtection = Data.define(
      :workbook_password, :lock_structure, :lock_windows, :revisions_password,
      :lock_revision, :algorithm_name, :hash_value, :salt_value, :spin_count
    ) do
      include MetadataRecord

      #: (?workbook_password: String?, ?lock_structure: bool?, ?lock_windows: bool?, ?revisions_password: String?, ?lock_revision: bool?, ?algorithm_name: String?, ?hash_value: String?, ?salt_value: String?, ?spin_count: Integer?, **untyped) -> void
      def initialize(workbook_password: nil, lock_structure: nil, lock_windows: nil, revisions_password: nil, lock_revision: nil, algorithm_name: nil, hash_value: nil, salt_value: nil, spin_count: nil, **_extra)
        super(
          workbook_password: workbook_password, lock_structure: lock_structure, lock_windows: lock_windows,
          revisions_password: revisions_password, lock_revision: lock_revision, algorithm_name: algorithm_name,
          hash_value: hash_value, salt_value: salt_value, spin_count: spin_count
        )
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
      :vertical_centered, :id
    ) do
      include MetadataRecord

      #: (?paper_size: (Integer | String)?, ?orientation: (Symbol | String)?, ?scale: Integer?, ?fit_to_width: Integer?, ?fit_to_height: Integer?, ?first_page_number: Integer?, ?use_first_page_number: bool?, ?black_and_white: bool?, ?draft: bool?, ?cell_comments: String?, ?use_printer_defaults: bool?, ?copies: Integer?, ?horizontal_centered: bool?, ?vertical_centered: bool?, ?id: String?, **untyped) -> void
      def initialize(paper_size: nil, orientation: nil, scale: nil, fit_to_width: nil, fit_to_height: nil, first_page_number: nil, use_first_page_number: nil, black_and_white: nil, draft: nil, cell_comments: nil, use_printer_defaults: nil, copies: nil, horizontal_centered: nil, vertical_centered: nil, id: nil, **_extra)
        super(
          paper_size: paper_size, orientation: orientation, scale: scale, fit_to_width: fit_to_width,
          fit_to_height: fit_to_height, first_page_number: first_page_number, use_first_page_number: use_first_page_number,
          black_and_white: black_and_white, draft: draft, cell_comments: cell_comments,
          use_printer_defaults: use_printer_defaults, copies: copies, horizontal_centered: horizontal_centered,
          vertical_centered: vertical_centered, id: id
        )
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

      #: (?left: Float?, ?right: Float?, ?top: Float?, ?bottom: Float?, ?header: Float?, ?footer: Float?, **untyped) -> void
      def initialize(left: nil, right: nil, top: nil, bottom: nil, header: nil, footer: nil, **_extra)
        super(left: left, right: right, top: top, bottom: bottom, header: header, footer: footer)
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

      #: (?odd_header: String?, ?odd_footer: String?, ?even_header: String?, ?even_footer: String?, ?first_header: String?, ?first_footer: String?, ?different_odd_even: bool?, ?different_first: bool?, ?scale_with_doc: bool?, ?align_with_margins: bool?, **untyped) -> void
      def initialize(odd_header: nil, odd_footer: nil, even_header: nil, even_footer: nil, first_header: nil, first_footer: nil, different_odd_even: nil, different_first: nil, scale_with_doc: nil, align_with_margins: nil, **_extra)
        super(
          odd_header: odd_header, odd_footer: odd_footer, even_header: even_header,
          even_footer: even_footer, first_header: first_header, first_footer: first_footer,
          different_odd_even: different_odd_even, different_first: different_first,
          scale_with_doc: scale_with_doc, align_with_margins: align_with_margins
        )
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

      #: (?grid_lines: bool?, ?grid_lines_set: bool?, ?headings: bool?, ?horizontal_centered: bool?, ?vertical_centered: bool?, **untyped) -> void
      def initialize(grid_lines: nil, grid_lines_set: nil, headings: nil, horizontal_centered: nil, vertical_centered: nil, **_extra)
        super(
          grid_lines: grid_lines, grid_lines_set: grid_lines_set, headings: headings,
          horizontal_centered: horizontal_centered, vertical_centered: vertical_centered
        )
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
    PrintTitles = Data.define(:rows, :columns, :cols) do
      include MetadataRecord

      #: (?rows: String?, ?columns: String?, ?cols: String?, **untyped) -> void
      def initialize(rows: nil, columns: nil, cols: nil, **_extra)
        c = columns || cols
        super(rows: rows, columns: c, cols: c)
      end

      undef_method :cols

      # @deprecated Use {#columns} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def cols
        warn "[Xlsxrb DEPRECATION] PrintTitles#cols is deprecated; use #columns instead.", uplevel: 1
        columns
      end
    end

    # Worksheet freeze pane settings (<pane>).
    # @api public
    FreezePane = Data.define(:row, :column, :col, :top_row, :left_column, :active_pane, :state) do
      include MetadataRecord

      #: (?row: Integer?, ?column: (Integer | String | Symbol)?, ?col: (Integer | String | Symbol)?, ?top_row: Integer?, ?left_column: (Integer | String | Symbol)?, ?active_pane: String?, ?state: Symbol?, **untyped) -> void
      def initialize(row: nil, column: nil, col: nil, top_row: nil, left_column: nil, active_pane: nil, state: :frozen, **_extra)
        c = column || col
        super(row: row, column: c, col: c, top_row: top_row, left_column: left_column, active_pane: active_pane, state: state || :frozen)
      end

      undef_method :col

      # @deprecated Use {#column} instead.
      # @return [Integer, String, Symbol, nil]
      # @api public
      #: () -> (Integer | String | Symbol)?
      def col
        warn "[Xlsxrb DEPRECATION] FreezePane#col is deprecated; use #column instead.", uplevel: 1
        column
      end
    end

    # Worksheet split pane settings (<pane>).
    # @api public
    SplitPane = Data.define(:x_split, :y_split, :top_row, :left_column, :active_pane, :state) do
      include MetadataRecord

      #: (?x_split: (Integer | Float)?, ?y_split: (Integer | Float)?, ?top_row: Integer?, ?left_column: Integer?, ?active_pane: String?, ?state: String?, **untyped) -> void
      def initialize(x_split: nil, y_split: nil, top_row: nil, left_column: nil, active_pane: nil, state: nil, **_extra)
        super(x_split: x_split, y_split: y_split, top_row: top_row, left_column: left_column, active_pane: active_pane, state: state)
      end
    end

    # Worksheet selection / active cell (<selection>).
    # @api public
    Selection = Data.define(:pane, :active_cell, :active_cell_id, :sqref) do
      include MetadataRecord

      #: (?pane: String?, ?active_cell: String?, ?active_cell_id: Integer?, ?sqref: String?, **untyped) -> void
      def initialize(pane: nil, active_cell: nil, active_cell_id: nil, sqref: nil, **_extra)
        super(pane: pane, active_cell: active_cell, active_cell_id: active_cell_id, sqref: sqref)
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

      #: (?password: String?, ?sheet: bool?, ?objects: bool?, ?scenarios: bool?, ?format_cells: bool?, ?format_columns: bool?, ?format_rows: bool?, ?insert_columns: bool?, ?insert_rows: bool?, ?insert_hyperlinks: bool?, ?delete_columns: bool?, ?delete_rows: bool?, ?select_locked_cells: bool?, ?select_unlocked_cells: bool?, ?sort: bool?, ?auto_filter: bool?, ?pivot_tables: bool?, ?algorithm_name: String?, ?hash_value: String?, ?salt_value: String?, ?spin_count: Integer?, **untyped) -> void
      def initialize(password: nil, sheet: nil, objects: nil, scenarios: nil, format_cells: nil, format_columns: nil, format_rows: nil, insert_columns: nil, insert_rows: nil, insert_hyperlinks: nil, delete_columns: nil, delete_rows: nil, select_locked_cells: nil, select_unlocked_cells: nil, sort: nil, auto_filter: nil, pivot_tables: nil, algorithm_name: nil, hash_value: nil, salt_value: nil, spin_count: nil, **_extra)
        super(
          password: password, sheet: sheet, objects: objects, scenarios: scenarios,
          format_cells: format_cells, format_columns: format_columns, format_rows: format_rows,
          insert_columns: insert_columns, insert_rows: insert_rows, insert_hyperlinks: insert_hyperlinks,
          delete_columns: delete_columns, delete_rows: delete_rows, select_locked_cells: select_locked_cells,
          select_unlocked_cells: select_unlocked_cells, sort: sort, auto_filter: auto_filter,
          pivot_tables: pivot_tables, algorithm_name: algorithm_name, hash_value: hash_value,
          salt_value: salt_value, spin_count: spin_count
        )
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
      :tab_color, :filter_mode, :enable_format_conditions_calculation, :outline_pr, :page_setup_properties,
      :code_name, :fit_to_page, :full_calc_on_load, :tab_color_theme, :tab_color_tint, :tab_color_indexed,
      :tab_color_auto, :outline_properties, :auto_page_breaks, :sync_horizontal, :sync_vertical, :sync_ref,
      :transition_evaluation, :transition_entry, :published,
      :apply_styles, :summary_below, :summary_right, :show_outline_symbols
    ) do
      include MetadataRecord

      #: (?tab_color: String?, ?filter_mode: bool?, ?enable_format_conditions_calculation: bool?, ?outline_pr: Hash[Symbol, untyped]?, ?page_setup_properties: Hash[Symbol, untyped]?, ?code_name: String?, ?fit_to_page: bool?, ?full_calc_on_load: bool?, ?tab_color_theme: Integer?, ?tab_color_tint: Float?, ?tab_color_indexed: Integer?, ?tab_color_auto: bool?, ?outline_properties: Hash[Symbol, untyped]?, ?auto_page_breaks: bool?, ?sync_horizontal: bool?, ?sync_vertical: bool?, ?sync_ref: String?, ?transition_evaluation: bool?, ?transition_entry: bool?, ?published: bool?, ?apply_styles: bool?, ?summary_below: bool?, ?summary_right: bool?, ?show_outline_symbols: bool?, **untyped) -> void
      def initialize(tab_color: nil, filter_mode: nil, enable_format_conditions_calculation: nil, outline_pr: nil, page_setup_properties: nil, page_setup_pr: nil, code_name: nil, fit_to_page: nil, full_calc_on_load: nil, tab_color_theme: nil, tab_color_tint: nil, tab_color_indexed: nil, tab_color_auto: nil, outline_properties: nil, auto_page_breaks: nil, sync_horizontal: nil, sync_vertical: nil, sync_ref: nil, transition_evaluation: nil, transition_entry: nil, published: nil, apply_styles: nil, summary_below: nil, summary_right: nil, show_outline_symbols: nil, **extra)
        psp = page_setup_properties
        if (page_setup_pr || extra.key?(:page_setup_pr) || extra.key?("page_setup_pr")) && psp.nil?
          Xlsxrb::Deprecation.warn("SheetProperties keyword :page_setup_pr is deprecated; use :page_setup_properties instead.", uplevel: 1)
          psp = page_setup_pr || extra[:page_setup_pr] || extra["page_setup_pr"]
        end
        out_pr = outline_properties || outline_pr || {}
        app_st = apply_styles.nil? ? (out_pr[:apply_styles] || extra[:apply_styles]) : apply_styles
        sum_b = summary_below.nil? ? (out_pr[:summary_below] || extra[:summary_below]) : summary_below
        sum_r = summary_right.nil? ? (out_pr[:summary_right] || extra[:summary_right]) : summary_right
        show_sym = show_outline_symbols.nil? ? (out_pr[:show_outline_symbols] || out_pr[:show_symbols] || extra[:show_outline_symbols] || extra[:show_symbols]) : show_outline_symbols
        super(
          tab_color: tab_color, filter_mode: filter_mode,
          enable_format_conditions_calculation: enable_format_conditions_calculation,
          outline_pr: outline_pr, page_setup_properties: psp,
          code_name: code_name, fit_to_page: fit_to_page, full_calc_on_load: full_calc_on_load,
          tab_color_theme: tab_color_theme, tab_color_tint: tab_color_tint, tab_color_indexed: tab_color_indexed,
          tab_color_auto: tab_color_auto, outline_properties: outline_properties, auto_page_breaks: auto_page_breaks,
          sync_horizontal: sync_horizontal, sync_vertical: sync_vertical, sync_ref: sync_ref,
          transition_evaluation: transition_evaluation, transition_entry: transition_entry, published: published,
          apply_styles: app_st, summary_below: sum_b, summary_right: sum_r, show_outline_symbols: show_sym
        )
      end

      # @deprecated Use {#page_setup_properties} instead. Will be removed in 0.2.0.
      def page_setup_pr
        Xlsxrb::Deprecation.warn("SheetProperties#page_setup_pr is deprecated; use SheetProperties#page_setup_properties instead.", uplevel: 1)
        page_setup_properties
      end

      undef_method :outline_pr

      # @deprecated Use {#outline_properties} instead.
      # @return [Hash{Symbol => untyped}, nil]
      # @api public
      #: () -> Hash[Symbol, untyped]?
      def outline_pr
        warn "[Xlsxrb DEPRECATION] SheetProperties#outline_pr is deprecated; use #outline_properties instead.", uplevel: 1
        outline_properties
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

      #: (?default_row_height: (Float | Integer)?, ?default_column_width: (Float | Integer)?, ?custom_height: bool?, ?zero_height: bool?, ?thick_top: bool?, ?thick_bottom: bool?, ?outline_level_row: Integer?, ?outline_level_column: Integer?, ?base_column_width: (Float | Integer)?, **untyped) -> void
      def initialize(default_row_height: nil, default_column_width: nil, custom_height: nil, zero_height: nil, thick_top: nil, thick_bottom: nil, outline_level_row: nil, outline_level_column: nil, base_column_width: nil, **extra)
        default_column_width ||= extra[:default_col_width]
        outline_level_column ||= extra[:outline_level_col]
        base_column_width ||= extra[:base_col_width]
        super(
          default_row_height: default_row_height, default_column_width: default_column_width,
          custom_height: custom_height, zero_height: zero_height, thick_top: thick_top,
          thick_bottom: thick_bottom, outline_level_row: outline_level_row,
          outline_level_column: outline_level_column, base_column_width: base_column_width
        )
      end

      # @deprecated Use {#default_column_width} instead.
      # @return [Float, Integer, nil]
      # @api public
      #: () -> (Float | Integer)?
      def default_col_width
        warn "[Xlsxrb DEPRECATION] SheetFormat#default_col_width is deprecated; use #default_column_width instead.", uplevel: 1
        default_column_width
      end

      # @deprecated Use {#outline_level_column} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def outline_level_col
        warn "[Xlsxrb DEPRECATION] SheetFormat#outline_level_col is deprecated; use #outline_level_column instead.", uplevel: 1
        outline_level_column
      end

      # @deprecated Use {#base_column_width} instead.
      # @return [Float, Integer, nil]
      # @api public
      #: () -> (Float | Integer)?
      def base_col_width
        warn "[Xlsxrb DEPRECATION] SheetFormat#base_col_width is deprecated; use #base_column_width instead.", uplevel: 1
        base_column_width
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

      #: (?workbook_view_id: Integer?, ?show_grid_lines: bool?, ?show_row_col_headers: bool?, ?show_ruler: bool?, ?show_zeros: bool?, ?right_to_left: bool?, ?tab_selected: bool?, ?zoom_scale: Integer?, ?zoom_scale_normal: Integer?, ?zoom_scale_page_layout_view: Integer?, ?zoom_scale_sheet_layout_view: Integer?, ?view: String?, ?window_protection: bool?, ?show_formulas: bool?, ?show_outline_symbols: bool?, ?default_grid_color: bool?, ?show_white_space: bool?, ?top_left_cell: String?, ?color_id: Integer?, **untyped) -> void
      def initialize(
        workbook_view_id: nil, show_grid_lines: nil, show_row_col_headers: nil, show_ruler: nil,
        show_zeros: nil, right_to_left: nil, tab_selected: nil, zoom_scale: nil, zoom_scale_normal: nil,
        zoom_scale_page_layout_view: nil, zoom_scale_sheet_layout_view: nil, view: nil,
        window_protection: nil, show_formulas: nil, show_outline_symbols: nil, default_grid_color: nil, show_white_space: nil,
        top_left_cell: nil, color_id: nil, **_extra
      )
        super(
          workbook_view_id: workbook_view_id, show_grid_lines: show_grid_lines,
          show_row_col_headers: show_row_col_headers, show_ruler: show_ruler,
          show_zeros: show_zeros, right_to_left: right_to_left, tab_selected: tab_selected,
          zoom_scale: zoom_scale, zoom_scale_normal: zoom_scale_normal,
          zoom_scale_page_layout_view: zoom_scale_page_layout_view,
          zoom_scale_sheet_layout_view: zoom_scale_sheet_layout_view, view: view,
          window_protection: window_protection, show_formulas: show_formulas,
          show_outline_symbols: show_outline_symbols, default_grid_color: default_grid_color,
          show_white_space: show_white_space, top_left_cell: top_left_cell, color_id: color_id
        )
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

      #: (?font_id: Integer?, ?type: String?, ?alignment: String?, **untyped) -> void
      def initialize(font_id: nil, type: nil, alignment: nil, **_extra)
        super(font_id: font_id, type: type, alignment: alignment)
      end
    end

    # Worksheet calculation properties (<sheetCalcPr>).
    # @api public
    SheetCalcProperties = Data.define(:full_calc_on_load) do
      include MetadataRecord

      #: (?full_calc_on_load: bool?, **untyped) -> void
      def initialize(full_calc_on_load: nil, **_extra)
        super(full_calc_on_load: full_calc_on_load)
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

      #: (?ref: String?, ?case_sensitive: bool?, ?sort_method: String?, ?conditions: Array[Hash[Symbol, untyped]]?, **untyped) -> void
      def initialize(ref: nil, case_sensitive: nil, sort_method: nil, conditions: nil, **_extra)
        super(ref: ref, case_sensitive: case_sensitive, sort_method: sort_method, conditions: conditions)
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
    BackgroundImage = Data.define(:r_id, :rel_id, :path, :extension, :data, :ext, :filename, :content_type) do
      include MetadataRecord

      #: (?r_id: String?, ?rel_id: String?, ?path: String?, ?extension: String?, ?data: String?, ?ext: String?, ?filename: String?, ?content_type: String?, **untyped) -> void
      def initialize(r_id: nil, rel_id: nil, path: nil, extension: nil, data: nil, ext: nil, filename: nil, content_type: nil, **extra)
        rid = r_id || rel_id
        e = extension || ext || extra[:ext]
        d = data || extra[:file_data]
        p = path || filename
        super(r_id: rid, rel_id: rid, path: p, extension: e, data: d, ext: e, filename: p, content_type: content_type)
      end

      undef_method :rel_id
      undef_method :ext
      undef_method :filename

      # @deprecated Use {#r_id} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def rel_id
        warn "[Xlsxrb DEPRECATION] BackgroundImage#rel_id is deprecated; use #r_id instead.", uplevel: 1
        r_id
      end

      # @deprecated Use {#extension} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def ext
        warn "[Xlsxrb DEPRECATION] BackgroundImage#ext is deprecated; use #extension instead.", uplevel: 1
        extension
      end

      # @deprecated Use {#path} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def filename
        warn "[Xlsxrb DEPRECATION] BackgroundImage#filename is deprecated; use #path instead.", uplevel: 1
        path
      end
    end

    # Table styles configuration in stylesheet (<tableStyles>).
    # @api public
    TableStyles = Data.define(:default_table_style, :default_pivot_style, :styles) do
      include MetadataRecord

      #: (?default_table_style: String?, ?default_pivot_style: String?, ?styles: Array[Hash[Symbol, untyped]]?, **untyped) -> void
      def initialize(default_table_style: nil, default_pivot_style: nil, styles: nil, **_extra)
        super(default_table_style: default_table_style, default_pivot_style: default_pivot_style, styles: styles)
      end
    end

    # Colors palette configuration in stylesheet (<colors>).
    # @api public
    Colors = Data.define(:indexed, :mru) do
      include MetadataRecord

      #: (?indexed: Array[String]?, ?mru: Array[Hash[Symbol, untyped]]?, **untyped) -> void
      def initialize(indexed: nil, mru: nil, **_extra)
        super(indexed: indexed, mru: mru)
      end
    end

    # User-defined custom document property (<property>).
    # @api public
    CustomProperty = Data.define(:name, :value, :type) do
      include MetadataRecord

      #: (?name: String?, ?value: untyped, ?type: Symbol?, **untyped) -> void
      def initialize(name: nil, value: nil, type: nil, **_extra)
        super(name: name, value: value, type: type)
      end

      # @return [String]
      # @api public
      #: () -> String
      def to_s
        value.to_s
      end

      # Matches against scalar values or hashes.
      #: (untyped other) -> bool
      def ==(other)
        if other.is_a?(CustomProperty)
          super
        elsif other.is_a?(Hash)
          warn "[Xlsxrb DEPRECATION] Comparing CustomProperty with Hash is deprecated; compare with CustomProperty instead.", uplevel: 1
          to_h == other.transform_keys(&:to_sym).compact
        else
          value == other
        end
      end
    end

    # AutoFilter column configuration (<filterColumn>).
    # @api public
    FilterColumn = Data.define(:column_id, :filters, :custom_filters, :dynamic_filter, :top10, :color_filter, :icon_filter, :hidden, :show_button) do
      include MetadataRecord

      #: (?column_id: Integer?, ?col_id: Integer?, ?filters: Array[untyped]?, ?custom_filters: Array[untyped]?, ?dynamic_filter: Hash[Symbol, untyped]?, ?top10: Hash[Symbol, untyped]?, ?color_filter: Hash[Symbol, untyped]?, ?icon_filter: Hash[Symbol, untyped]?, ?hidden: bool?, ?show_button: bool?, ?hidden_button: bool?, **untyped) -> void
      def initialize(column_id: nil, col_id: nil, filters: nil, custom_filters: nil, dynamic_filter: nil, top10: nil, color_filter: nil, icon_filter: nil, hidden: false, show_button: true, hidden_button: nil, **extra)
        if (!col_id.nil? || extra.key?(:col_id)) && column_id.nil?
          warn "[Xlsxrb DEPRECATION] FilterColumn.new(col_id:) is deprecated and will be removed in v0.2.0. Use column_id: instead.", uplevel: 1
          column_id = col_id || extra[:col_id]
        end
        h = hidden || hidden_button || false
        super(column_id: column_id, filters: filters, custom_filters: custom_filters, dynamic_filter: dynamic_filter, top10: top10, color_filter: color_filter, icon_filter: icon_filter, hidden: h, show_button: show_button)
      end

      # @deprecated Use {#column_id} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def col_id
        warn "[Xlsxrb DEPRECATION] FilterColumn#col_id is deprecated and will be removed in v0.2.0. Use #column_id instead.", uplevel: 1
        column_id
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
    end

    # Workbook defined name (<definedName>).
    # @api public
    DefinedName = Data.define(:name, :value, :local_sheet_id, :local_sheet_name, :hidden, :comment) do
      include MetadataRecord

      #: (?name: String?, ?value: String?, ?local_sheet_id: Integer?, ?local_sheet_name: String?, ?hidden: bool?, ?comment: String?, **untyped) -> void
      def initialize(name: nil, value: nil, local_sheet_id: nil, local_sheet_name: nil, hidden: nil, comment: nil, **extra)
        local_sheet_name ||= extra[:sheet]
        value ||= extra[:formula]
        if (value.nil? || (value.respond_to?(:empty?) && value.empty?)) && extra[:range]
          value = local_sheet_name ? "'#{local_sheet_name}'!#{extra[:range]}" : extra[:range].to_s
        end
        super(name: name, value: value || "", local_sheet_id: local_sheet_id, local_sheet_name: local_sheet_name, hidden: hidden, comment: comment)
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

      #: (?id: Integer?, ?max: Integer?, ?min: Integer?, ?man: (Integer | bool)?, ?pt: bool?, ?manual: bool?, **untyped) -> void
      # rubocop:disable Naming/MethodParameterName
      def initialize(id: nil, max: nil, min: nil, man: nil, pt: nil, manual: nil, **_extra)
        m = if manual.nil?
              man
            else
              manual ? 1 : 0
            end
        super(id: id, max: max, min: min, man: m, pt: pt)
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
    # @api public
    Comment = Data.define(:ref, :text, :author, :r_id, :shape_id, :author_id, :visible) do
      include MetadataRecord

      #: (?ref: String?, ?cell: String?, ?text: String?, ?author: String?, ?r_id: String?, ?shape_id: Integer?, ?author_id: Integer?, ?visible: bool?, **untyped) -> void
      def initialize(ref: nil, cell: nil, text: nil, author: nil, r_id: nil, shape_id: nil, author_id: nil, visible: false, **_extra)
        r = ref || cell
        super(ref: r, text: text, author: author, r_id: r_id, shape_id: shape_id, author_id: author_id, visible: visible)
      end

      # @deprecated Use {#ref} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def cell
        warn "[Xlsxrb DEPRECATION] Comment#cell is deprecated; use #ref instead.", uplevel: 1
        ref
      end

      # Returns whether the comment is visible by default.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def visible?
        visible ? true : false
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
    ThreadedComment = Data.define(:id, :ref, :text, :author, :author_id, :created_at, :parent_id, :resolved) do
      include MetadataRecord

      #: (?id: String?, ?ref: String?, ?cell: String?, ?text: String?, ?author: String?, ?author_id: String?, ?person_id: String?, ?created_at: untyped, ?timestamp: untyped, ?dt: untyped, ?parent_id: String?, ?resolved: bool?, ?done: bool?, **untyped) -> void
      # rubocop:disable Naming/MethodParameterName
      def initialize(
        id: nil, ref: nil, cell: nil, text: nil, author: nil,
        author_id: nil, person_id: nil, created_at: nil, timestamp: nil, dt: nil,
        parent_id: nil, resolved: nil, done: nil, **_extra
      )
        r = ref || cell
        res = resolved.nil? ? done : resolved
        aid = author_id || person_id
        cat = created_at || timestamp || dt
        super(id: id, ref: r, text: text, author: author, author_id: aid, created_at: cat, parent_id: parent_id, resolved: res)
      end
      # rubocop:enable Naming/MethodParameterName

      # @deprecated Use {#ref} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def cell
        warn "[Xlsxrb DEPRECATION] ThreadedComment#cell is deprecated; use #ref instead.", uplevel: 1
        ref
      end

      # @deprecated Use {#resolved} instead.
      # @return [Boolean, nil]
      # @api public
      #: () -> bool?
      def done
        warn "[Xlsxrb DEPRECATION] ThreadedComment#done is deprecated; use #resolved instead.", uplevel: 1
        resolved
      end

      # Returns whether the threaded comment is resolved.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def resolved?
        resolved ? true : false
      end

      # @deprecated Use {#resolved?} instead.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def done?
        warn "[Xlsxrb DEPRECATION] ThreadedComment#done? is deprecated; use #resolved? instead.", uplevel: 1
        resolved?
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

      #: (?sqref: String?, ?type: String?, ?error_style: String?, ?operator: String?, ?allow_blank: bool?, ?show_input_message: bool?, ?show_error_message: bool?, ?prompt_title: String?, ?prompt: String?, ?error_title: String?, ?error: String?, ?formula1: String?, ?formula2: String?, ?ime_mode: String?, ?show_drop_down: bool?, **untyped) -> void
      def initialize(sqref: nil, type: nil, error_style: nil, operator: nil, allow_blank: nil, show_input_message: nil, show_error_message: nil, prompt_title: nil, prompt: nil, error_title: nil, error: nil, formula1: nil, formula2: nil, ime_mode: nil, show_drop_down: nil, **extra)
        sdd = show_drop_down.nil? ? extra[:show_dropdown] : show_drop_down
        super(
          sqref: sqref, type: type, error_style: error_style, operator: operator,
          allow_blank: allow_blank, show_input_message: show_input_message,
          show_error_message: show_error_message, prompt_title: prompt_title, prompt: prompt,
          error_title: error_title, error: error, formula1: formula1, formula2: formula2,
          ime_mode: ime_mode, show_drop_down: sdd
        )
      end

      # @deprecated Use {#show_drop_down} instead.
      # @return [Boolean, nil]
      # @api public
      #: () -> bool?
      def show_dropdown
        warn "[Xlsxrb DEPRECATION] DataValidation#show_dropdown is deprecated; use #show_drop_down instead.", uplevel: 1
        show_drop_down
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

      # @deprecated Use {#show_drop_down?} instead.
      # @return [Boolean]
      # @api public
      #: () -> bool
      def show_dropdown?
        warn "[Xlsxrb DEPRECATION] DataValidation#show_dropdown? is deprecated; use #show_drop_down? instead.", uplevel: 1
        show_drop_down?
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

      #: (?type: (Symbol | String)?, ?dxf_id: Integer?, ?priority: Integer?, ?operator: (Symbol | String)?, ?formula: untyped, ?text: String?, ?time_period: (Symbol | String)?, ?rank: Integer?, ?std_dev: Integer?, ?above_average: bool?, ?equal_average: bool?, ?bottom: bool?, ?percent: bool?, ?stop_if_true: bool?, ?color_scale: Hash[Symbol, untyped]?, ?data_bar: Hash[Symbol, untyped]?, ?icon_set: Hash[Symbol, untyped]?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(type: nil, dxf_id: nil, priority: nil, operator: nil, formula: nil, text: nil, time_period: nil, rank: nil, std_dev: nil, above_average: nil, equal_average: nil, bottom: nil, percent: nil, stop_if_true: nil, color_scale: nil, data_bar: nil, icon_set: nil, options: nil, **extra)
        dxf_id ||= extra.delete(:format_id)
        f = formula || extra.delete(:formulas)
        opts = (options || {}).merge(extra)
        super(
          type: type, dxf_id: dxf_id, priority: priority, operator: operator, formula: f,
          text: text, time_period: time_period, rank: rank, std_dev: std_dev,
          above_average: above_average, equal_average: equal_average, bottom: bottom,
          percent: percent, stop_if_true: stop_if_true, color_scale: color_scale,
          data_bar: data_bar, icon_set: icon_set, options: opts.empty? ? nil : opts
        )
      end

      # @deprecated Use {#formula} instead.
      # @return [Object]
      # @api public
      #: () -> untyped
      def formulas
        warn "[Xlsxrb DEPRECATION] ConditionalFormatRule#formulas is deprecated; use #formula instead.", uplevel: 1
        formula
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

      #: (?sqref: String?, ?type: (Symbol | String)?, ?dxf_id: Integer?, ?priority: Integer?, ?operator: (Symbol | String)?, ?formula: untyped, ?text: String?, ?time_period: (Symbol | String)?, ?rank: Integer?, ?std_dev: Integer?, ?above_average: bool?, ?equal_average: bool?, ?bottom: bool?, ?percent: bool?, ?stop_if_true: bool?, ?color_scale: Hash[Symbol, untyped]?, ?data_bar: Hash[Symbol, untyped]?, ?icon_set: Hash[Symbol, untyped]?, ?pivot: bool?, ?rules: Array[untyped]?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(sqref: nil, type: nil, dxf_id: nil, priority: nil, operator: nil, formula: nil, text: nil, time_period: nil, rank: nil, std_dev: nil, above_average: nil, equal_average: nil, bottom: nil, percent: nil, stop_if_true: nil, color_scale: nil, data_bar: nil, icon_set: nil, pivot: nil, rules: nil, options: nil, **extra)
        dxf_id ||= extra.delete(:format_id)
        f = formula || extra.delete(:formulas)
        wrapped_rules = if rules && !rules.empty?
                          rules.map { |r| r.is_a?(ConditionalFormatRule) ? r : ConditionalFormatRule.new(**r) }
                        end
        opts = (options || {}).merge(extra)
        super(
          sqref: sqref, type: type, dxf_id: dxf_id, priority: priority, operator: operator, formula: f,
          text: text, time_period: time_period, rank: rank, std_dev: std_dev,
          above_average: above_average, equal_average: equal_average, bottom: bottom,
          percent: percent, stop_if_true: stop_if_true, color_scale: color_scale,
          data_bar: data_bar, icon_set: icon_set, pivot: pivot, rules: wrapped_rules,
          options: opts.empty? ? nil : opts
        )
      end

      # @deprecated Use {#formula} instead.
      # @return [Object]
      # @api public
      #: () -> untyped
      def formulas
        warn "[Xlsxrb DEPRECATION] ConditionalFormat#formulas is deprecated; use #formula instead.", uplevel: 1
        formula
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
    TableColumn = Data.define(:id, :name, :totals_row_function, :totals_row_label, :totals_row_formula, :data_dxf_id, :calculated_column_formula) do
      include MetadataRecord

      #: (?id: Integer?, ?name: String?, ?totals_row_function: String?, ?totals_row_label: String?, ?totals_row_formula: String?, ?data_dxf_id: Integer?, ?calculated_column_formula: String?, ?data_formula: String?, **untyped) -> void
      def initialize(id: nil, name: nil, totals_row_function: nil, totals_row_label: nil, totals_row_formula: nil, data_dxf_id: nil, calculated_column_formula: nil, data_formula: nil, **_extra)
        super(
          id: id, name: name, totals_row_function: totals_row_function,
          totals_row_label: totals_row_label, totals_row_formula: totals_row_formula,
          data_dxf_id: data_dxf_id,
          calculated_column_formula: calculated_column_formula || data_formula
        )
      end

      # @deprecated Use {#calculated_column_formula} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def data_formula
        warn "[Xlsxrb DEPRECATION] TableColumn#data_formula is deprecated; use #calculated_column_formula instead.", uplevel: 1
        calculated_column_formula
      end
    end

    # Table style information (<tableStyleInfo>).
    # @api public
    TableStyleInfo = Data.define(:name, :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes) do
      include MetadataRecord

      #: (?name: String?, ?show_first_column: bool?, ?show_last_column: bool?, ?show_row_stripes: bool?, ?show_column_stripes: bool?, **untyped) -> void
      def initialize(name: nil, show_first_column: nil, show_last_column: nil, show_row_stripes: nil, show_column_stripes: nil, **_extra)
        super(
          name: name, show_first_column: show_first_column, show_last_column: show_last_column,
          show_row_stripes: show_row_stripes, show_column_stripes: show_column_stripes
        )
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
        super(
          id: id, name: name, display_name: display_name, ref: ref,
          totals_row_count: trc, header_row_count: hrc,
          columns: wrapped_cols, style_info: wrapped_si, auto_filter: auto_filter
        )
      end

      #: (Symbol | String key) -> untyped
      def [](key)
        warn "[DEPRECATION] Hash access on Elements::Table (`#[]`) is deprecated and will be removed in 0.2.0; use method calls instead.", uplevel: 1
        k = key.is_a?(Symbol) ? key : key.to_s.to_sym
        case k
        when :style
          style_info&.name
        when :show_row_stripes
          style_info&.show_row_stripes
        when :show_first_column
          style_info&.show_first_column
        when :show_last_column
          style_info&.show_last_column
        when :show_column_stripes
          style_info&.show_column_stripes
        else
          super
        end
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
      alias_method :show_row_stripes, :show_row_stripes?

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

      #: (?sqref: String?, ?formula: String?, **untyped) -> void
      def initialize(sqref: nil, formula: nil, **extra)
        sqref ||= extra[:location_ref] || extra[:location]
        formula ||= extra[:data_ref] || extra[:f] || extra[:data]
        super(sqref: sqref, formula: formula)
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

      #: (?type: String?, ?color_series: String?, ?color_negative: String?, ?color_axis: String?, ?color_markers: String?, ?color_first: String?, ?color_last: String?, ?color_high: String?, ?color_low: String?, ?sparklines: Array[Elements::Sparkline | Hash[Symbol, untyped]]?, ?display_empty_cells_as: String?, ?markers: bool?, ?high: bool?, ?low: bool?, ?first: bool?, ?last: bool?, ?negative: bool?, ?manual_max: (Integer | Float)?, ?manual_min: (Integer | Float)?, ?line_weight: (Integer | Float)?, **untyped) -> void
      def initialize(type: nil, color_series: nil, color_negative: nil, color_axis: nil, color_markers: nil, color_first: nil, color_last: nil, color_high: nil, color_low: nil, sparklines: nil, display_empty_cells_as: nil, markers: nil, high: nil, low: nil, first: nil, last: nil, negative: nil, manual_max: nil, manual_min: nil, line_weight: nil, **extra)
        wrapped_sps = (sparklines || []).map { |s| s.is_a?(Sparkline) ? s : Sparkline.new(**s) }
        disp = display_empty_cells_as || extra[:display_empty]
        m_max = manual_max || extra[:max]
        m_min = manual_min || extra[:min]
        lw = line_weight || extra[:line_weight]
        super(
          type: type, color_series: color_series, color_negative: color_negative,
          color_axis: color_axis, color_markers: color_markers, color_first: color_first,
          color_last: color_last, color_high: color_high, color_low: color_low,
          sparklines: wrapped_sps, display_empty_cells_as: disp,
          markers: markers, high: high, low: low, first: first, last: last, negative: negative,
          manual_max: m_max, manual_min: m_min, line_weight: lw
        )
      end

      # @deprecated Use {#display_empty_cells_as} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def display_empty
        warn "[Xlsxrb DEPRECATION] SparklineGroup#display_empty is deprecated; use #display_empty_cells_as instead.", uplevel: 1
        display_empty_cells_as
      end

      # @deprecated Use {#manual_max} instead.
      # @return [Integer, Float, nil]
      # @api public
      #: () -> (Integer | Float)?
      def max
        warn "[Xlsxrb DEPRECATION] SparklineGroup#max is deprecated; use #manual_max instead.", uplevel: 1
        manual_max
      end

      # @deprecated Use {#manual_min} instead.
      # @return [Integer, Float, nil]
      # @api public
      #: () -> (Integer | Float)?
      def min
        warn "[Xlsxrb DEPRECATION] SparklineGroup#min is deprecated; use #manual_min instead.", uplevel: 1
        manual_min
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

      #: (?name: String?, ?sqref: String?, ?security_descriptor: String?, ?password: String?, ?algorithm_name: String?, ?hash_value: String?, ?salt_value: String?, ?spin_count: Integer?, **untyped) -> void
      def initialize(name: nil, sqref: nil, security_descriptor: nil, password: nil, algorithm_name: nil, hash_value: nil, salt_value: nil, spin_count: nil, **_extra)
        super(
          name: name, sqref: sqref, security_descriptor: security_descriptor, password: password,
          algorithm_name: algorithm_name, hash_value: hash_value, salt_value: salt_value, spin_count: spin_count
        )
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

      #: (?sqref: String?, ?eval_error: bool?, ?two_digit_text_year: bool?, ?number_stored_as_text: bool?, ?formula: bool?, ?formula_range: bool?, ?unlocked_formula: bool?, ?empty_cell_reference: bool?, ?list_data_validation: bool?, ?calculated_column: bool?, **untyped) -> void
      def initialize(sqref: nil, eval_error: nil, two_digit_text_year: nil, number_stored_as_text: nil, formula: nil, formula_range: nil, unlocked_formula: nil, empty_cell_reference: nil, list_data_validation: nil, calculated_column: nil, **_extra)
        super(
          sqref: sqref, eval_error: eval_error, two_digit_text_year: two_digit_text_year,
          number_stored_as_text: number_stored_as_text, formula: formula, formula_range: formula_range,
          unlocked_formula: unlocked_formula, empty_cell_reference: empty_cell_reference,
          list_data_validation: list_data_validation, calculated_column: calculated_column
        )
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

      #: (?r_id: String?, ?name: String?, ?title: String?, ?type: String?, ?anchor: String?, ?from: untyped, ?to: untyped, ?series: Array[untyped]?, ?legend: untyped, ?category_axis: untyped, ?value_axis: untyped, ?secondary_value_axis: untyped, ?grouping: untyped, ?gap_width: untyped, ?overlap: untyped, ?target: untyped, ?from_column: untyped, ?from_row: untyped, ?to_column: untyped, ?to_row: untyped, ?data_labels: untyped?, ?cat_axis_title: String?, ?val_axis_title: String?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(r_id: nil, name: nil, title: nil, type: nil, anchor: nil, from: nil, to: nil, series: nil, legend: nil, category_axis: nil, value_axis: nil, secondary_value_axis: nil, grouping: nil, gap_width: nil, overlap: nil, target: nil, from_column: nil, from_row: nil, to_column: nil, to_row: nil, data_labels: nil, cat_axis_title: nil, val_axis_title: nil, options: nil, **extra)
        from_column ||= extra.delete(:from_col)
        to_column ||= extra.delete(:to_col)
        r_id ||= extra.delete(:rid)
        type ||= extra.delete(:chart_type)
        cat_title = cat_axis_title || extra.delete(:cat_axis_title)
        val_title = val_axis_title || extra.delete(:val_axis_title)
        category_axis ||= { title: cat_title } if cat_title
        value_axis ||= { title: val_title } if val_title
        opts = (options || {}).merge(extra)
        super(
          r_id: r_id, name: name, title: title, type: type, anchor: anchor, from: from, to: to,
          series: series, legend: legend, category_axis: category_axis, value_axis: value_axis,
          secondary_value_axis: secondary_value_axis, grouping: grouping, gap_width: gap_width,
          overlap: overlap, target: target, from_column: from_column, from_row: from_row,
          to_column: to_column, to_row: to_row, data_labels: data_labels,
          options: opts.empty? ? nil : opts
        )
      end

      # @deprecated Use {#category_axis}[:title] instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def cat_axis_title
        warn "[Xlsxrb DEPRECATION] Chart#cat_axis_title is deprecated and will be removed in v0.2.0. Use #category_axis[:title] instead.", uplevel: 1
        category_axis&.[](:title) || options&.[](:cat_axis_title)
      end

      # @deprecated Use {#value_axis}[:title] instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def val_axis_title
        warn "[Xlsxrb DEPRECATION] Chart#val_axis_title is deprecated and will be removed in v0.2.0. Use #value_axis[:title] instead.", uplevel: 1
        value_axis&.[](:title) || options&.[](:val_axis_title)
      end

      # @deprecated Use {#r_id} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def rid
        warn "[Xlsxrb DEPRECATION] Chart#rid is deprecated; use #r_id instead.", uplevel: 1
        r_id
      end

      # @deprecated Use {#type} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def chart_type
        warn "[Xlsxrb DEPRECATION] Chart#chart_type is deprecated; use #type instead.", uplevel: 1
        type
      end

      # @deprecated Use {#from_column} instead.
      # @return [Object]
      # @api public
      #: () -> untyped
      def from_col
        warn "[Xlsxrb DEPRECATION] Chart#from_col is deprecated; use #from_column instead.", uplevel: 1
        from_column
      end

      # @deprecated Use {#to_column} instead.
      # @return [Object]
      # @api public
      #: () -> untyped
      def to_col
        warn "[Xlsxrb DEPRECATION] Chart#to_col is deprecated; use #to_column instead.", uplevel: 1
        to_column
      end
    end

    # Drawing shape entry (<xdr:sp>).
    # @api public
    Shape = Data.define(:id, :name, :type, :anchor, :from, :to, :text, :preset, :width, :height, :left, :top, :options) do
      include MetadataRecord

      #: (?id: Integer?, ?name: String?, ?type: String?, ?anchor: String?, ?from: untyped, ?to: untyped, ?text: String?, ?preset: String?, ?width: (Float | Integer)?, ?height: (Float | Integer)?, ?left: (Float | Integer)?, ?top: (Float | Integer)?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(id: nil, name: nil, type: nil, anchor: nil, from: nil, to: nil, text: nil, preset: nil, width: nil, height: nil, left: nil, top: nil, options: nil, **extra)
        opts = (options || {}).merge(extra)
        super(
          id: id, name: name, type: type, anchor: anchor, from: from, to: to,
          text: text, preset: preset, width: width, height: height, left: left, top: top,
          options: opts.empty? ? nil : opts
        )
      end
    end

    # Pivot table entry (<pivotTableDefinition>).
    # @api public
    PivotTable = Data.define(
      :name, :ref, :location, :source_ref, :row_fields, :column_fields,
      :data_fields, :page_fields, :data_caption, :page_wrap,
      :column_grand_totals, :row_grand_totals, :show_row_headers,
      :show_column_headers, :show_row_stripes, :show_column_stripes,
      :cache, :options
    ) do
      include MetadataRecord

      #: (?name: String?, ?ref: String?, ?location: String?, ?source_ref: String?, ?row_fields: Array[untyped]?, ?column_fields: Array[untyped]?, ?col_fields: Array[untyped]?, ?data_fields: Array[untyped]?, ?page_fields: Array[untyped]?, ?data_caption: String?, ?page_wrap: Integer?, ?column_grand_totals: bool?, ?row_grand_totals: bool?, ?show_row_headers: bool?, ?show_column_headers: bool?, ?show_row_stripes: bool?, ?show_column_stripes: bool?, ?cache: Hash[Symbol, untyped]?, ?options: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(
        name: nil, ref: nil, location: nil, source_ref: nil,
        row_fields: nil, column_fields: nil, col_fields: nil, data_fields: nil, page_fields: nil,
        data_caption: nil, page_wrap: nil, column_grand_totals: nil, row_grand_totals: nil,
        show_row_headers: nil, show_column_headers: nil, show_row_stripes: nil, show_column_stripes: nil,
        cache: nil, options: nil, **extra
      )
        r = ref || location
        cols = column_fields || col_fields || []
        merged_opts = (options || {}).merge(extra).freeze
        super(
          name: name, ref: r, location: r, source_ref: source_ref,
          row_fields: row_fields, column_fields: cols, data_fields: data_fields, page_fields: page_fields,
          data_caption: data_caption, page_wrap: page_wrap,
          column_grand_totals: column_grand_totals, row_grand_totals: row_grand_totals,
          show_row_headers: show_row_headers, show_column_headers: show_column_headers,
          show_row_stripes: show_row_stripes, show_column_stripes: show_column_stripes,
          cache: cache, options: merged_opts.empty? ? nil : merged_opts
        )
      end

      undef_method :location

      # @deprecated Use {#ref} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def location
        warn "[Xlsxrb DEPRECATION] PivotTable#location is deprecated; use #ref instead.", uplevel: 1
        ref
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

      #: (?position: String?, ?r_id: String?, ?path: String?, ?extension: String?, ?data: String?, ?content_type: String?, **untyped) -> void
      def initialize(position: nil, r_id: nil, path: nil, extension: nil, data: nil, content_type: nil, **extra)
        rid = r_id || extra[:rid]
        ext = extension || extra[:ext]
        p = path || extra[:filename]
        ct = content_type || extra[:content_type]
        super(position: position, r_id: rid, path: p, extension: ext, data: data, content_type: ct)
      end

      # @deprecated Use {#extension} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def ext
        warn "[Xlsxrb DEPRECATION] HeaderFooterImage#ext is deprecated; use #extension instead.", uplevel: 1
        extension
      end

      # @deprecated Use {#path} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def filename
        warn "[Xlsxrb DEPRECATION] HeaderFooterImage#filename is deprecated; use #path instead.", uplevel: 1
        path
      end

      # @deprecated Use {#r_id} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def rid
        warn "[Xlsxrb DEPRECATION] HeaderFooterImage#rid is deprecated; use #r_id instead.", uplevel: 1
        r_id
      end
    end

    # Font style definition in stylesheet (<font>).
    # @api public
    # rubocop:disable Lint/DataDefineOverride
    Font = Data.define(:name, :size, :bold, :italic, :underline, :strike, :color, :family, :scheme, :outline, :shadow, :condense, :extend) do
      # rubocop:enable Lint/DataDefineOverride
      include MetadataRecord

      #: (?name: String?, ?size: (Integer | Float)?, ?bold: bool?, ?italic: bool?, ?underline: (String | bool)?, ?strike: bool?, ?color: Hash[Symbol, untyped]?, ?family: Integer?, ?scheme: String?, ?outline: bool?, ?shadow: bool?, ?condense: bool?, ?extend: bool?, **untyped) -> void
      def initialize(name: nil, size: nil, bold: nil, italic: nil, underline: nil, strike: nil, color: nil, family: nil, scheme: nil, outline: nil, shadow: nil, condense: nil, extend: nil, **extra)
        sz = size || extra[:sz]
        super(name: name, size: sz, bold: bold, italic: italic, underline: underline, strike: strike, color: color, family: family, scheme: scheme, outline: outline, shadow: shadow, condense: condense, extend: extend)
      end

      # @deprecated Use {#size} instead.
      # @return [Integer, Float, nil]
      # @api public
      #: () -> (Integer | Float)?
      def sz
        warn "[Xlsxrb DEPRECATION] Font#sz is deprecated; use #size instead.", uplevel: 1
        size
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
    # @api public
    Fill = Data.define(:pattern_type, :fg_color, :bg_color) do
      include MetadataRecord

      #: (?pattern_type: String?, ?fg_color: Hash[Symbol, untyped]?, ?bg_color: Hash[Symbol, untyped]?, **untyped) -> void
      def initialize(pattern_type: nil, fg_color: nil, bg_color: nil, **extra)
        pt = pattern_type || extra[:pattern]
        super(pattern_type: pt, fg_color: fg_color, bg_color: bg_color)
      end

      # @deprecated Use {#pattern_type} instead.
      # @return [String, nil]
      # @api public
      #: () -> String?
      def pattern
        warn "[Xlsxrb DEPRECATION] Fill#pattern is deprecated; use #pattern_type instead.", uplevel: 1
        pattern_type
      end
    end

    # Border style definition in stylesheet (<border>).
    # @api public
    Border = Data.define(:left, :right, :top, :bottom, :diagonal, :diagonal_up, :diagonal_down) do
      include MetadataRecord

      #: (?left: Hash[Symbol, untyped]?, ?right: Hash[Symbol, untyped]?, ?top: Hash[Symbol, untyped]?, ?bottom: Hash[Symbol, untyped]?, ?diagonal: Hash[Symbol, untyped]?, ?diagonal_up: bool?, ?diagonal_down: bool?, **untyped) -> void
      def initialize(left: nil, right: nil, top: nil, bottom: nil, diagonal: nil, diagonal_up: nil, diagonal_down: nil, **_extra)
        super(left: left, right: right, top: top, bottom: bottom, diagonal: diagonal, diagonal_up: diagonal_up, diagonal_down: diagonal_down)
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

      #: (?horizontal: (String | Symbol)?, ?vertical: (String | Symbol)?, ?wrap_text: bool?, ?text_rotation: Numeric?, ?indent: Integer?, ?relative_indent: Integer?, ?shrink_to_fit: bool?, ?reading_order: Integer?, ?justify_last_line: bool?, **untyped) -> void
      def initialize(horizontal: nil, vertical: nil, wrap_text: nil, text_rotation: nil, indent: nil, relative_indent: nil, shrink_to_fit: nil, reading_order: nil, justify_last_line: nil, **_extra)
        super(horizontal: horizontal, vertical: vertical, wrap_text: wrap_text, text_rotation: text_rotation, indent: indent, relative_indent: relative_indent, shrink_to_fit: shrink_to_fit, reading_order: reading_order, justify_last_line: justify_last_line)
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

      #: (?locked: bool?, ?hidden: bool?, **untyped) -> void
      def initialize(locked: nil, hidden: nil, **_extra)
        super(locked: locked, hidden: hidden)
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

      #: (?num_fmt_id: Integer?, ?font_id: Integer?, ?fill_id: Integer?, ?border_id: Integer?, ?xf_id: Integer?, ?apply_number_format: bool?, ?apply_font: bool?, ?apply_fill: bool?, ?apply_border: bool?, ?apply_alignment: bool?, ?apply_protection: bool?, ?alignment: (Elements::Alignment | Hash[Symbol, untyped])?, ?protection: (Elements::Protection | Hash[Symbol, untyped])?, ?quote_prefix: bool?, ?pivot_button: bool?, **untyped) -> void
      def initialize(
        num_fmt_id: nil, font_id: nil, fill_id: nil, border_id: nil, xf_id: nil,
        apply_number_format: nil, apply_font: nil, apply_fill: nil, apply_border: nil,
        apply_alignment: nil, apply_protection: nil, alignment: nil, protection: nil,
        quote_prefix: nil, pivot_button: nil, **_extra
      )
        align_obj = alignment.is_a?(Hash) ? Alignment.new(**alignment) : alignment
        prot_obj = protection.is_a?(Hash) ? Protection.new(**protection) : protection
        super(
          num_fmt_id: num_fmt_id, font_id: font_id, fill_id: fill_id, border_id: border_id,
          xf_id: xf_id, apply_number_format: apply_number_format, apply_font: apply_font,
          apply_fill: apply_fill, apply_border: apply_border, apply_alignment: apply_alignment,
          apply_protection: apply_protection, alignment: align_obj, protection: prot_obj,
          quote_prefix: quote_prefix, pivot_button: pivot_button
        )
      end

      # @deprecated Use {#num_fmt_id} instead.
      # @return [Integer, nil]
      # @api public
      #: () -> Integer?
      def number_format_id
        warn "[Xlsxrb DEPRECATION] CellXf#number_format_id is deprecated; use #num_fmt_id instead.", uplevel: 1
        num_fmt_id
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

      #: (?font: (Elements::Font | Hash[Symbol, untyped])?, ?fill: (Elements::Fill | Hash[Symbol, untyped])?, ?border: (Elements::Border | Hash[Symbol, untyped])?, ?number_format: Hash[Symbol, untyped]?, ?alignment: (Elements::Alignment | Hash[Symbol, untyped])?, ?protection: (Elements::Protection | Hash[Symbol, untyped])?, **untyped) -> void
      def initialize(font: nil, fill: nil, border: nil, number_format: nil, alignment: nil, protection: nil, **extra)
        nf = number_format
        if (extra.key?(:num_fmt) || extra.key?("num_fmt")) && nf.nil?
          Xlsxrb::Deprecation.warn("Dxf keyword :num_fmt is deprecated; use :number_format instead.", uplevel: 1)
          nf = extra[:num_fmt] || extra["num_fmt"]
        end
        f = font.is_a?(Hash) ? Font.new(**font) : font
        fi = fill.is_a?(Hash) ? Fill.new(**fill) : fill
        b = border.is_a?(Hash) ? Border.new(**border) : border
        align = alignment.is_a?(Hash) ? Alignment.new(**alignment) : alignment
        prot = protection.is_a?(Hash) ? Protection.new(**protection) : protection
        super(font: f, fill: fi, border: b, number_format: nf, alignment: align, protection: prot)
      end

      # @deprecated Use {#number_format} instead. Will be removed in 0.2.0.
      def num_fmt
        Xlsxrb::Deprecation.warn("Dxf#num_fmt is deprecated; use Dxf#number_format instead.", uplevel: 1)
        number_format
      end
    end

    # Named cell style in stylesheet (<cellStyle> in <cellStyles>).
    # @api public
    CellStyle = Data.define(:name, :xf_id, :builtin_id) do
      include MetadataRecord

      #: (?name: String?, ?xf_id: Integer?, ?builtin_id: Integer?, **untyped) -> void
      def initialize(name: nil, xf_id: nil, builtin_id: nil, **_extra)
        super(name: name, xf_id: xf_id, builtin_id: builtin_id)
      end
    end

    # Color definition in stylesheet (<color>).
    # @api public
    Color = Data.define(:rgb, :indexed, :theme, :tint) do
      include MetadataRecord

      #: (?rgb: String?, ?indexed: Integer?, ?theme: Integer?, ?tint: Float?, **untyped) -> void
      def initialize(rgb: nil, indexed: nil, theme: nil, tint: nil, **_extra)
        super(rgb: rgb, indexed: indexed, theme: theme, tint: tint)
      end
    end
  end
end
