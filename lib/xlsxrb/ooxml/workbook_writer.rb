# frozen_string_literal: true

require "stringio"
require_relative "xml_builder"
require_relative "zip_writer"
require_relative "worksheet_writer"

module Xlsxrb
  module Ooxml
    # Orchestrates writing a complete XLSX workbook.
    # Generates all required OpenXML parts and assembles them into a ZIP.
    class WorkbookWriter
      SSML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      REL_NS  = "http://schemas.openxmlformats.org/package/2006/relationships"
      CT_NS   = "http://schemas.openxmlformats.org/package/2006/content-types"
      DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

      def self.write(target, sheets:, shared_strings: [], styles: nil,
                     defined_names: nil, core_properties: nil, app_properties: nil,
                     custom_properties: nil, workbook_protection: nil)
        writer = new(sheets: sheets, shared_strings: shared_strings, styles: styles,
                     defined_names: defined_names, core_properties: core_properties,
                     app_properties: app_properties, custom_properties: custom_properties,
                     workbook_protection: workbook_protection)
        writer.write_to(target)
      end

      def initialize(sheets:, shared_strings: [], styles: nil,
                     defined_names: nil, core_properties: nil, app_properties: nil,
                     custom_properties: nil, workbook_protection: nil)
        @sheets = sheets
        @shared_strings = shared_strings
        @styles = styles
        @defined_names = defined_names || []
        @core_properties = core_properties || {}
        @app_properties = app_properties || {}
        @custom_properties = custom_properties || []
        @workbook_protection = workbook_protection
        @drawing_count = 0
        @chart_count = 0
        @comment_count = 0
        @table_count = 0
      end

      def write_to(target)
        ZipWriter.open(target) do |zip|
          zip.add_entry("[Content_Types].xml", build_content_types)
          zip.add_entry("_rels/.rels", build_root_rels)
          zip.add_entry("xl/workbook.xml", build_workbook_xml)
          zip.add_entry("xl/_rels/workbook.xml.rels", build_workbook_rels)
          zip.add_entry("xl/styles.xml", build_styles_xml)
          zip.add_entry("xl/sharedStrings.xml", build_shared_strings_xml) unless @shared_strings.empty?

          # Document properties
          zip.add_entry("docProps/core.xml", build_core_properties_xml) unless @core_properties.empty?
          zip.add_entry("docProps/app.xml", build_app_properties_xml) unless @app_properties.empty?
          zip.add_entry("docProps/custom.xml", build_custom_properties_xml) unless @custom_properties.empty?

          @sheets.each_with_index do |sheet, idx|
            sheet_images = sheet[:images] || []
            sheet_charts = sheet[:charts] || []
            sheet_shapes = sheet[:shapes] || []
            sheet_comments = sheet[:comments] || []
            sheet_tables = sheet[:tables] || []
            sheet_hyperlinks = sheet[:hyperlinks] || []
            has_drawing = sheet_images.any? || sheet_charts.any? || sheet_shapes.any?
            has_comments = sheet_comments.any?

            # Track relationship IDs for this sheet
            sheet_rels = []
            drawing_rid = nil
            comment_rid = nil
            vml_rid = nil
            table_start_rid = nil
            hyperlink_rels = []

            # Drawing relationships (images, charts, shapes)
            if has_drawing
              @drawing_count += 1
              drawing_rid_num = sheet_rels.size + 1
              sheet_rels << { id: "rId#{drawing_rid_num}", type: "#{DOC_REL}/drawing", target: "../drawings/drawing#{@drawing_count}.xml" }
              drawing_rid = "rId#{drawing_rid_num}"

              drawing_rels_data = []
              drawing_parts = []

              chart_writer = Xlsxrb::Ooxml::Writer.new
              chart_writer.add_sheet(sheet[:name])

              # Populate sheet data in chart_writer for chart cache resolution
              (sheet[:rows] || []).each do |row|
                (row[:cells] || []).each do |cell|
                  chart_writer.set_cell(cell[:ref], cell[:value], sheet: sheet[:name])
                end
              end

              sheet_images.each do |img|
                media_idx = @drawing_count # simplified
                media_path = "xl/media/image#{media_idx}.#{img[:ext] || 'png'}"
                zip.add_binary_entry(media_path, img[:file_data])
                drawing_rels_data << { type: :image, target: "../media/image#{media_idx}.#{img[:ext] || 'png'}" }
                drawing_parts << { kind: :pic, img: img, rid_index: drawing_rels_data.size }
              end

              sheet_charts.each do |chart_options|
                chart_writer.add_chart(**chart_options)
              end

              processed_charts = chart_writer.charts
              processed_charts.each do |chart|
                @chart_count += 1
                chart_path = "xl/charts/chart#{@chart_count}.xml"
                zip.add_entry(chart_path, chart_writer.send(:generate_chart_xml, chart))
                drawing_rels_data << { type: :chart, target: "../charts/chart#{@chart_count}.xml" }
                drawing_parts << { kind: :chart, chart: chart, rid_index: drawing_rels_data.size }
              end

              sheet_shapes.each_with_index do |shape, si|
                drawing_parts << { kind: :sp, shape: shape, id: drawing_parts.size + si + 2 }
              end

              drawing_xml = chart_writer.send(:generate_drawing_xml, drawing_parts)
              zip.add_entry("xl/drawings/drawing#{@drawing_count}.xml", drawing_xml)
              unless drawing_rels_data.empty?
                drawing_rels_xml = chart_writer.send(:generate_drawing_rels, drawing_rels_data)
                zip.add_entry("xl/drawings/_rels/drawing#{@drawing_count}.xml.rels", drawing_rels_xml)
              end
            end

            # Comment relationships
            if has_comments
              @comment_count += 1
              comment_rid_num = sheet_rels.size + 1
              sheet_rels << { id: "rId#{comment_rid_num}", type: "#{DOC_REL}/comments", target: "../comments#{@comment_count}.xml" }
              comment_rid = "rId#{comment_rid_num}"
              vml_rid_num = sheet_rels.size + 1
              sheet_rels << { id: "rId#{vml_rid_num}", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", target: "../drawings/vmlDrawing#{@comment_count}.vml" }
              vml_rid = "rId#{vml_rid_num}"

              # Generate comments XML using Writer's method
              comment_writer = Xlsxrb::Ooxml::Writer.new
              sheet_comments.each do |c|
                comment_writer.add_comment(c[:cell], c[:text], author: c[:author] || "Author")
              end
              zip.add_entry("xl/comments#{@comment_count}.xml", comment_writer.send(:generate_comments_xml, comment_writer.comments))
              zip.add_entry("xl/drawings/vmlDrawing#{@comment_count}.vml", comment_writer.send(:generate_vml_drawing_xml, comment_writer.comments))
            end

            # Hyperlink relationships (external URLs need rels)
            sheet_hyperlinks.each_with_index do |link, hi|
              next unless link[:url]

              h_rid_num = sheet_rels.size + 1
              sheet_rels << { id: "rId#{h_rid_num}", type: "#{DOC_REL}/hyperlink", target: link[:url], target_mode: "External" }
              hyperlink_rels << h_rid_num
            end
            # Assign rIds to hyperlinks
            h_rid_idx = 0
            enriched_hyperlinks = sheet_hyperlinks.map do |link|
              if link[:url]
                rid = hyperlink_rels[h_rid_idx]
                h_rid_idx += 1
                link.merge(_rid: rid)
              else
                link
              end
            end

            # Table relationships
            unless sheet_tables.empty?
              table_start_rid = sheet_rels.size + 1
              sheet_tables.each_with_index do |_tbl, ti|
                @table_count += 1
                t_rid_num = sheet_rels.size + 1
                sheet_rels << { id: "rId#{t_rid_num}", type: "#{DOC_REL}/table", target: "../tables/table#{@table_count}.xml" }
              end
            end

            # Build worksheet rels if any
            unless sheet_rels.empty?
              zip.add_entry("xl/worksheets/_rels/sheet#{idx + 1}.xml.rels", build_sheet_rels_from_list(sheet_rels))
            end

            # Generate table XML files
            table_id_base = @table_count - sheet_tables.size
            sheet_tables.each_with_index do |tbl, ti|
              tbl_id = table_id_base + ti + 1
              zip.add_entry("xl/tables/table#{tbl_id}.xml", build_table_xml(tbl, tbl_id))
            end

            # Build the worksheet XML with all metadata
            zip.add_entry("xl/worksheets/sheet#{idx + 1}.xml", build_worksheet_xml(
              sheet,
              drawing_rid: drawing_rid,
              sheet_protection: sheet[:sheet_protection],
              auto_filter: sheet[:auto_filter],
              filter_columns: sheet[:filter_columns],
              sort_state: sheet[:sort_state],
              merge_cells: sheet[:merge_cells],
              conditional_formats: sheet[:conditional_formats],
              data_validations: sheet[:data_validations],
              hyperlinks: enriched_hyperlinks.empty? ? nil : enriched_hyperlinks,
              print_options: sheet[:print_options],
              page_margins: sheet[:page_margins],
              page_setup: sheet[:page_setup],
              header_footer: sheet[:header_footer],
              row_breaks: sheet[:row_breaks],
              col_breaks: sheet[:col_breaks],
              freeze_pane: sheet[:freeze_pane],
              split_pane: sheet[:split_pane],
              selection: sheet[:selection],
              sheet_view: sheet[:sheet_view],
              sheet_properties: sheet[:sheet_properties],
              tables: sheet_tables,
              table_start_rid: table_start_rid,
              legacy_drawing_rid: vml_rid
            ))
          end
        end
      end

      private

      def build_sheet_rels_from_list(rels)
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Relationships", { xmlns: REL_NS })
        rels.each do |rel|
          attrs = { Id: rel[:id], Type: rel[:type], Target: rel[:target] }
          attrs[:TargetMode] = rel[:target_mode] if rel[:target_mode]
          b.empty_tag("Relationship", attrs)
        end
        b.close_tag("Relationships")
        io.string
      end

      def build_content_types
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Types", { xmlns: CT_NS })
        b.empty_tag("Default", { Extension: "rels", ContentType: "application/vnd.openxmlformats-package.relationships+xml" })
        b.empty_tag("Default", { Extension: "xml", ContentType: "application/xml" })
        b.empty_tag("Default", { Extension: "vml", ContentType: "application/vnd.openxmlformats-officedocument.vmlDrawing" })
        b.empty_tag("Override", { PartName: "/xl/workbook.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" })
        b.empty_tag("Override", { PartName: "/xl/styles.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" })
        b.empty_tag("Override", { PartName: "/xl/sharedStrings.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" }) unless @shared_strings.empty?

        # Document properties
        b.empty_tag("Override", { PartName: "/docProps/core.xml", ContentType: "application/vnd.openxmlformats-package.core-properties+xml" }) unless @core_properties.empty?
        b.empty_tag("Override", { PartName: "/docProps/app.xml", ContentType: "application/vnd.openxmlformats-officedocument.extended-properties+xml" }) unless @app_properties.empty?
        b.empty_tag("Override", { PartName: "/docProps/custom.xml", ContentType: "application/vnd.openxmlformats-officedocument.custom-properties+xml" }) unless @custom_properties.empty?

        drawing_count = 0
        chart_count = 0
        comment_count = 0
        table_count = 0
        image_exts = {}
        @sheets.each_with_index do |sheet, idx|
          b.empty_tag("Override", { PartName: "/xl/worksheets/sheet#{idx + 1}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" })

          sheet_images = sheet[:images] || []
          sheet_charts = sheet[:charts] || []
          sheet_shapes = sheet[:shapes] || []
          sheet_comments = sheet[:comments] || []
          sheet_tables = sheet[:tables] || []
          has_drawing = sheet_images.any? || sheet_charts.any? || sheet_shapes.any?

          if has_drawing
            drawing_count += 1
            b.empty_tag("Override", { PartName: "/xl/drawings/drawing#{drawing_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.drawing+xml" })
            sheet_charts.each do |_chart|
              chart_count += 1
              b.empty_tag("Override", { PartName: "/xl/charts/chart#{chart_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" })
            end
            sheet_images.each do |img|
              ext = img[:ext] || "png"
              unless image_exts[ext]
                mime = case ext
                       when "png" then "image/png"
                       when "jpg", "jpeg" then "image/jpeg"
                       when "gif" then "image/gif"
                       when "bmp" then "image/bmp"
                       when "emf" then "image/x-emf"
                       when "wmf" then "image/x-wmf"
                       else "application/octet-stream"
                       end
                b.empty_tag("Default", { Extension: ext, ContentType: mime })
                image_exts[ext] = true
              end
            end
          end

          if sheet_comments.any?
            comment_count += 1
            b.empty_tag("Override", { PartName: "/xl/comments#{comment_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml" })
          end

          sheet_tables.each do |_tbl|
            table_count += 1
            b.empty_tag("Override", { PartName: "/xl/tables/table#{table_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" })
          end
        end

        b.close_tag("Types")
        io.string
      end

      def build_root_rels
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Relationships", { xmlns: REL_NS })
        b.empty_tag("Relationship", {
                      Id: "rId1",
                      Type: "#{DOC_REL}/officeDocument",
                      Target: "xl/workbook.xml"
                    })
        rid = 1
        unless @core_properties.empty?
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                        Target: "docProps/core.xml"
                      })
        end
        unless @app_properties.empty?
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                        Target: "docProps/app.xml"
                      })
        end
        unless @custom_properties.empty?
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties",
                        Target: "docProps/custom.xml"
                      })
        end
        b.close_tag("Relationships")
        io.string
      end

      def build_workbook_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("workbook", {
                     xmlns: SSML_NS,
                     "xmlns:r": DOC_REL
                   })

        # Workbook protection
        if @workbook_protection
          wp_attrs = {}
          wp_attrs[:lockStructure] = "1" if @workbook_protection[:lock_structure]
          wp_attrs[:lockWindows] = "1" if @workbook_protection[:lock_windows]
          wp_attrs[:workbookPassword] = @workbook_protection[:password] if @workbook_protection[:password]
          b.empty_tag("workbookProtection", wp_attrs) unless wp_attrs.empty?
        end

        b.open_tag("sheets")
        @sheets.each_with_index do |sheet, idx|
          b.empty_tag("sheet", {
                        name: sheet[:name],
                        sheetId: (idx + 1).to_s,
                        "r:id": "rId#{idx + 1}"
                      })
        end
        b.close_tag("sheets")

        # Defined names
        unless @defined_names.empty?
          b.open_tag("definedNames")
          @defined_names.each do |dn|
            dn_attrs = { name: dn[:name] }
            dn_attrs[:localSheetId] = dn[:local_sheet_id].to_s if dn[:local_sheet_id]
            dn_attrs[:hidden] = "1" if dn[:hidden]
            b.tag("definedName", dn_attrs) { |_| b.text(dn[:value]) }
          end
          b.close_tag("definedNames")
        end

        b.close_tag("workbook")
        io.string
      end

      def build_workbook_rels
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Relationships", { xmlns: REL_NS })
        @sheets.each_with_index do |_sheet, idx|
          b.empty_tag("Relationship", {
                        Id: "rId#{idx + 1}",
                        Type: "#{DOC_REL}/worksheet",
                        Target: "worksheets/sheet#{idx + 1}.xml"
                      })
        end
        rid = @sheets.size + 1
        b.empty_tag("Relationship", {
                      Id: "rId#{rid}",
                      Type: "#{DOC_REL}/styles",
                      Target: "styles.xml"
                    })
        unless @shared_strings.empty?
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "#{DOC_REL}/sharedStrings",
                        Target: "sharedStrings.xml"
                      })
        end
        b.close_tag("Relationships")
        io.string
      end

      def build_styles_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("styleSheet", { xmlns: SSML_NS })

        # Fonts
        fonts = @styles&.dig(:fonts) || []
        fonts = [{}] if fonts.empty? # At least one default font
        b.tag("fonts", { count: fonts.size.to_s }) do |_|
          fonts.each do |font_props|
            b.tag("font") do |_|
              b.empty_tag("sz", { val: (font_props[:sz] || 11).to_s }) if font_props[:sz] || fonts == [{}]
              b.empty_tag("b") if font_props[:bold]
              b.empty_tag("i") if font_props[:italic]
              b.empty_tag("strike") if font_props[:strike]
              if font_props[:underline]
                b.empty_tag("u", font_props[:underline] == "single" ? {} : { val: font_props[:underline] })
              end
              b.empty_tag("color", { rgb: font_props[:color] }) if font_props[:color]
              b.empty_tag("name", { val: font_props[:name] || "Calibri" })
            end
          end
        end

        # Fills
        fills = @styles&.dig(:fills) || []
        fills = [{ pattern: "none" }, { pattern: "gray125" }] if fills.empty?
        b.tag("fills", { count: fills.size.to_s }) do |_|
          fills.each do |fill_props|
            b.tag("fill") do |_|
              if fill_props[:gradient]
                b.open_tag("gradientFill", { type: fill_props[:gradient][:type], degree: fill_props[:gradient][:degree]&.to_s }.compact)
                fill_props[:gradient][:stops].each do |stop|
                  b.empty_tag("stop", { position: stop[:position].to_s })
                  b.empty_tag("color", { rgb: stop[:color] })
                end
                b.close_tag("gradientFill")
              else
                b.tag("patternFill", { patternType: fill_props[:pattern] || "none" }) do |_|
                  b.empty_tag("fgColor", { rgb: fill_props[:fg_color] }) if fill_props[:fg_color]
                  b.empty_tag("bgColor", { rgb: fill_props[:bg_color] }) if fill_props[:bg_color]
                end
              end
            end
          end
        end

        # Borders
        borders = @styles&.dig(:borders) || []
        borders = [{}] if borders.empty?
        b.tag("borders", { count: borders.size.to_s }) do |_|
          borders.each do |border_props|
            border_attrs = {}
            border_attrs[:diagonalUp] = "1" if border_props[:diagonal_up]
            border_attrs[:diagonalDown] = "1" if border_props[:diagonal_down]
            b.tag("border", border_attrs) do |_|
              %i[left right top bottom diagonal].each do |side|
                side_data = border_props[side]
                if side_data
                  b.tag(side.to_s, { style: side_data[:style] }) do |_|
                    b.empty_tag("color", { rgb: side_data[:color] }) if side_data[:color]
                  end
                else
                  b.empty_tag(side.to_s)
                end
              end
            end
          end
        end

        # cellStyleXfs
        b.tag("cellStyleXfs", { count: "1" }) do |_|
          b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" })
        end

        # Number formats
        num_fmts = @styles&.dig(:num_fmts) || []
        unless num_fmts.empty?
          b.tag("numFmts", { count: num_fmts.size.to_s }) do |_|
            num_fmts.each do |nf|
              if nf.is_a?(Hash)
                b.empty_tag("numFmt", { numFmtId: nf[:num_fmt_id].to_s, formatCode: nf[:format_code] })
              else
                # Legacy format handling
              end
            end
          end
        end

        # Cell XFs (formatting)
        xf_entries = @styles&.dig(:xf_entries) || []
        xf_count = [xf_entries.size, 1].max
        b.tag("cellXfs", { count: xf_count.to_s }) do |_|
          xf_entries.each do |xf|
            attrs = {
              numFmtId: (xf[:num_fmt_id] || 0).to_s,
              fontId: (xf[:font_id] || 0).to_s,
              fillId: (xf[:fill_id] || 0).to_s,
              borderId: (xf[:border_id] || 0).to_s
            }
            attrs[:xfId] = xf[:xf_id].to_s if xf[:xf_id]
            b.empty_tag("xf", attrs)
          end
          # Add default if empty
          b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" }) if xf_entries.empty?
        end

        b.close_tag("styleSheet")
        io.string
      end

      def build_shared_strings_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("sst", {
                     xmlns: SSML_NS,
                     count: @shared_strings.size.to_s,
                     uniqueCount: @shared_strings.size.to_s
                   })
        @shared_strings.each do |str|
          b.tag("si") { |_| b.tag("t") { |_| b.text(str) } }
        end
        b.close_tag("sst")
        io.string
      end

      def build_worksheet_xml(sheet, drawing_rid: nil, sheet_protection: nil,
                              auto_filter: nil, filter_columns: nil, sort_state: nil,
                              merge_cells: nil, conditional_formats: nil,
                              data_validations: nil, hyperlinks: nil,
                              print_options: nil, page_margins: nil, page_setup: nil,
                              header_footer: nil, row_breaks: nil, col_breaks: nil,
                              freeze_pane: nil, split_pane: nil, selection: nil,
                              sheet_view: nil, sheet_properties: nil,
                              tables: nil, table_start_rid: nil,
                              legacy_drawing_rid: nil)
        io = StringIO.new
        ws = WorksheetWriter.new(io)
        ws.start(
          columns: sheet[:columns] || [],
          sheet_properties: sheet_properties,
          freeze_pane: freeze_pane,
          split_pane: split_pane,
          selection: selection,
          sheet_view: sheet_view
        )
        (sheet[:rows] || []).each do |row|
          ws.write_row(row[:index], row[:cells], attrs: row[:attrs] || {}, unmapped: row[:unmapped] || [])
        end
        ws.finish(
          drawing_rid: drawing_rid,
          sheet_protection: sheet_protection,
          auto_filter: auto_filter,
          filter_columns: filter_columns,
          sort_state: sort_state,
          merge_cells: merge_cells,
          conditional_formats: conditional_formats,
          data_validations: data_validations,
          hyperlinks: hyperlinks,
          print_options: print_options,
          page_margins: page_margins,
          page_setup: page_setup,
          header_footer: header_footer,
          row_breaks: row_breaks,
          col_breaks: col_breaks,
          tables: tables,
          table_start_rid: table_start_rid,
          legacy_drawing_rid: legacy_drawing_rid
        )
        io.string
      end

      def build_table_xml(tbl, table_id)
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        tbl_name = tbl[:name] || "Table#{table_id}"
        t_attrs = {
          xmlns: SSML_NS,
          id: table_id.to_s,
          name: tbl_name,
          displayName: tbl[:display_name] || tbl_name,
          ref: tbl[:ref]
        }
        t_attrs[:totalsRowCount] = tbl[:totals_row_count].to_s if tbl[:totals_row_count] && tbl[:totals_row_count] > 0
        b.open_tag("table", t_attrs)

        # Auto filter for the table range
        b.empty_tag("autoFilter", { ref: tbl[:ref] })

        # Table columns
        columns = tbl[:columns] || []
        b.open_tag("tableColumns", { count: columns.size.to_s })
        columns.each_with_index do |col_name, ci|
          b.empty_tag("tableColumn", { id: (ci + 1).to_s, name: col_name })
        end
        b.close_tag("tableColumns")

        # Table style
        style = tbl[:style] || {}
        style_attrs = { name: style[:name] || "TableStyleMedium2" }
        style_attrs[:showFirstColumn] = style[:show_first_column] ? "1" : "0" unless style[:show_first_column].nil?
        style_attrs[:showLastColumn] = style[:show_last_column] ? "1" : "0" unless style[:show_last_column].nil?
        style_attrs[:showRowStripes] = style[:show_row_stripes] ? "1" : "0" unless style[:show_row_stripes].nil?
        style_attrs[:showColumnStripes] = style[:show_column_stripes] ? "1" : "0" unless style[:show_column_stripes].nil?
        b.empty_tag("tableStyleInfo", style_attrs)

        b.close_tag("table")
        io.string
      end

      def build_core_properties_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("cp:coreProperties", {
                     "xmlns:cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
                     "xmlns:dc": "http://purl.org/dc/elements/1.1/",
                     "xmlns:dcterms": "http://purl.org/dc/terms/",
                     "xmlns:dcmitype": "http://purl.org/dc/dcmitype/",
                     "xmlns:xsi": "http://www.w3.org/2001/XMLSchema-instance"
                   })
        b.tag("dc:title") { |_| b.text(@core_properties[:title]) } if @core_properties[:title]
        b.tag("dc:subject") { |_| b.text(@core_properties[:subject]) } if @core_properties[:subject]
        b.tag("dc:creator") { |_| b.text(@core_properties[:creator]) } if @core_properties[:creator]
        b.tag("cp:keywords") { |_| b.text(@core_properties[:keywords]) } if @core_properties[:keywords]
        b.tag("dc:description") { |_| b.text(@core_properties[:description]) } if @core_properties[:description]
        b.tag("cp:lastModifiedBy") { |_| b.text(@core_properties[:last_modified_by]) } if @core_properties[:last_modified_by]
        b.tag("cp:category") { |_| b.text(@core_properties[:category]) } if @core_properties[:category]
        b.close_tag("cp:coreProperties")
        io.string
      end

      def build_app_properties_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Properties", {
                     xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
                     "xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                   })
        b.tag("Application") { |_| b.text(@app_properties[:application]) } if @app_properties[:application]
        b.tag("Company") { |_| b.text(@app_properties[:company]) } if @app_properties[:company]
        b.tag("Manager") { |_| b.text(@app_properties[:manager]) } if @app_properties[:manager]
        b.close_tag("Properties")
        io.string
      end

      def build_custom_properties_xml
        io = StringIO.new
        b = XmlBuilder.new(io)
        b.declaration
        b.open_tag("Properties", {
                     xmlns: "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties",
                     "xmlns:vt": "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
                   })
        @custom_properties.each_with_index do |prop, idx|
          b.open_tag("property", { fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", pid: (idx + 2).to_s, name: prop[:name] })
          case prop[:type]
          when :number
            b.tag("vt:i4") { |_| b.text(prop[:value].to_s) }
          when :bool
            b.tag("vt:bool") { |_| b.text(prop[:value] ? "true" : "false") }
          when :date
            b.tag("vt:filetime") { |_| b.text(prop[:value].to_s) }
          else
            b.tag("vt:lpwstr") { |_| b.text(prop[:value].to_s) }
          end
          b.close_tag("property")
        end
        b.close_tag("Properties")
        io.string
      end
    end
  end
end
