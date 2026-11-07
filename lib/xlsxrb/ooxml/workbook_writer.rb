# frozen_string_literal: true

# rbs_inline: enabled

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

      # Wraps a ZipWriter entry to provide a standard IO-like #write interface.
      class ZipEntryIO
        def initialize(zip)
          @zip = zip
        end

        def write(data)
          @zip.write_data(data)
        end

        alias << write
      end

      def self.write(target, sheets:, shared_strings: [], shared_strings_index: nil, styles: nil,
                     defined_names: nil, core_properties: nil, app_properties: nil,
                     custom_properties: nil, workbook_protection: nil, workbook_properties: nil,
                     compression_level: Zlib::DEFAULT_COMPRESSION)
        writer = new(sheets: sheets, shared_strings: shared_strings, shared_strings_index: shared_strings_index, styles: styles,
                     defined_names: defined_names, core_properties: core_properties,
                     app_properties: app_properties, custom_properties: custom_properties,
                     workbook_protection: workbook_protection, workbook_properties: workbook_properties,
                     compression_level: compression_level)
        writer.write_to(target)
      end

      def initialize(sheets:, shared_strings: [], shared_strings_index: nil, styles: nil,
                     defined_names: nil, core_properties: nil, app_properties: nil,
                     custom_properties: nil, workbook_protection: nil, workbook_properties: nil,
                     compression_level: Zlib::DEFAULT_COMPRESSION)
        @sheets = sheets
        @shared_strings = shared_strings || []
        @shared_strings_index = shared_strings_index
        @styles = styles || {}
        @defined_names = defined_names || []
        @core_properties = core_properties.is_a?(Elements::MetadataRecord) ? core_properties.to_h : (core_properties || {})
        @app_properties = app_properties.is_a?(Elements::MetadataRecord) ? app_properties.to_h : (app_properties || {})
        @custom_properties = custom_properties || []
        @workbook_protection = workbook_protection.is_a?(Elements::MetadataRecord) ? workbook_protection.to_h : workbook_protection
        @workbook_properties = workbook_properties.is_a?(Elements::MetadataRecord) ? workbook_properties.to_h : (workbook_properties || { update_links: "never" })
        @compression_level = compression_level || Zlib::DEFAULT_COMPRESSION
        @drawing_count = 0
        @chart_count = 0
        @comment_count = 0
        @table_count = 0
        @pivot_cache_count = 0
        @pivot_table_count = 0
        @threaded_comment_count = 0
        @persons = []
      end

      def write_to(target)
        preprocess_conditional_formats!

        active_tab_idx = resolve_active_tab_index
        effective_active_idx = active_tab_idx || 0

        ZipWriter.open(target, compression_level: @compression_level) do |zip|
          @sheets.each_with_index do |sheet, idx|
            info = prepare_sheet_auxiliary(sheet, idx)
            sheet_tables = sheet[:tables] || []

            sv = sheet[:sheet_view]
            sv = sv.to_h if sv && !sv.is_a?(Hash)
            if idx == effective_active_idx
              sv = (sv || {}).merge(tab_selected: true)
            elsif sv&.key?(:tab_selected)
              sv = sv.dup
              sv.delete(:tab_selected)
            end

            zip.start_entry("xl/worksheets/sheet#{idx + 1}.xml")
            write_worksheet_xml(
              ZipEntryIO.new(zip),
              sheet,
              drawing_rid: info[:drawing_rid],
              picture_rid: info[:picture_rid],
              sheet_protection: sheet[:sheet_protection],
              auto_filter: sheet[:auto_filter],
              filter_columns: sheet[:filter_columns],
              sort_state: sheet[:sort_state],
              merge_cells: sheet[:merge_cells],
              conditional_formats: sheet[:conditional_formats],
              data_validations: sheet[:data_validations],
              hyperlinks: info[:enriched_hyperlinks].empty? ? nil : info[:enriched_hyperlinks],
              print_options: sheet[:print_options],
              page_margins: sheet[:page_margins],
              page_setup: sheet[:page_setup],
              header_footer: sheet[:header_footer],
              row_breaks: sheet[:row_breaks],
              column_breaks: sheet[:column_breaks],
              freeze_pane: sheet[:freeze_pane],
              split_pane: sheet[:split_pane],
              selection: sheet[:selection],
              sheet_view: sv,
              sheet_properties: sheet[:sheet_properties],
              sheet_format: sheet[:sheet_format],
              protected_ranges: sheet[:protected_ranges],
              cell_watches: sheet[:cell_watches],
              ignored_errors: sheet[:ignored_errors],
              phonetic_pr: sheet[:phonetic_properties] || sheet[:phonetic_pr],
              sheet_calc_pr: sheet[:sheet_calc_pr],
              tables: sheet_tables,
              table_start_rid: info[:table_start_rid],
              legacy_drawing_rid: info[:vml_rid],
              sparkline_groups: sheet[:sparkline_groups]
            )
            zip.finish_entry
          end

          write_package_parts(zip)
        end
      end

      def prepare_sheet_auxiliary(sheet, idx)
        sheet_images = sheet[:images] || []
        sheet_charts = sheet[:charts] || []
        sheet_shapes = sheet[:shapes] || []
        sheet_comments = sheet[:comments] || []
        sheet_threaded_comments = sheet[:threaded_comments] || []
        sheet_tables = sheet[:tables] || []
        sheet_hyperlinks = sheet[:hyperlinks] || []
        has_drawing = sheet_images.any? || sheet_charts.any? || sheet_shapes.any?
        has_comments = sheet_comments.any?
        has_threaded_comments = sheet_threaded_comments.any?

        sheet_rels = []
        drawing_rid = nil
        vml_rid = nil
        table_start_rid = nil
        picture_rid = nil
        hyperlink_rels = []

        if has_drawing
          @drawing_count += 1
          drawing_rid_num = sheet_rels.size + 1
          sheet_rels << { id: "rId#{drawing_rid_num}", type: "#{DOC_REL}/drawing", target: "../drawings/drawing#{@drawing_count}.xml" }
          drawing_rid = "rId#{drawing_rid_num}"
        end

        if has_comments
          @comment_count += 1
          comment_rid_num = sheet_rels.size + 1
          sheet_rels << { id: "rId#{comment_rid_num}", type: "#{DOC_REL}/comments", target: "../comments#{@comment_count}.xml" }
          vml_rid_num = sheet_rels.size + 1
          sheet_rels << { id: "rId#{vml_rid_num}", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing", target: "../drawings/vmlDrawing#{@comment_count}.vml" }
          vml_rid = "rId#{vml_rid_num}"
        end

        if has_threaded_comments
          @threaded_comment_count += 1
          tc_rid_num = sheet_rels.size + 1
          sheet_rels << { id: "rId#{tc_rid_num}", type: "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment", target: "../threadedComments/threadedComment#{@threaded_comment_count}.xml" }
        end

        sheet_hyperlinks.each do |link|
          next unless link[:url]

          h_rid_num = sheet_rels.size + 1
          sheet_rels << { id: "rId#{h_rid_num}", type: "#{DOC_REL}/hyperlink", target: link[:url], target_mode: "External" }
          hyperlink_rels << h_rid_num
        end

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

        unless sheet_tables.empty?
          table_start_rid = sheet_rels.size + 1
          sheet_tables.each do
            @table_count += 1
            t_rid_num = sheet_rels.size + 1
            sheet_rels << { id: "rId#{t_rid_num}", type: "#{DOC_REL}/table", target: "../tables/table#{@table_count}.xml" }
          end
        end

        sheet_pivot_tables = sheet[:pivot_tables] || []
        unless sheet_pivot_tables.empty?
          sheet_pivot_tables.each do
            @pivot_cache_count += 1
            @pivot_table_count += 1
            pt_rid_num = sheet_rels.size + 1
            sheet_rels << { id: "rId#{pt_rid_num}", type: "#{DOC_REL}/pivotTable", target: "../pivotTables/pivotTable#{@pivot_table_count}.xml" }
          end
        end

        bg_img = sheet[:background_image] || sheet[:background_picture] || sheet[:picture]
        if bg_img
          bg_data = if bg_img.is_a?(Hash)
                      bg_img[:file_data] || bg_img[:data]
                    else
                      (bg_img.respond_to?(:data) ? bg_img.data : bg_img)
                    end
          if bg_data.is_a?(String) && !bg_data.include?("\x00") && File.file?(bg_data)
            bg_ext = File.extname(bg_data).delete_prefix(".").downcase
            bg_data = File.binread(bg_data)
          else
            bg_ext = if bg_img.is_a?(Hash)
                       (bg_img[:ext] || (bg_img[:filename] && File.extname(bg_img[:filename]).delete_prefix(".")) || "png").to_s.downcase
                     elsif bg_img.respond_to?(:filename)
                       File.extname(bg_img.filename).delete_prefix(".").downcase
                     else
                       "png"
                     end
          end
          if bg_data && !bg_data.empty?
            p_rid_num = sheet_rels.size + 1
            picture_rid = "rId#{p_rid_num}"
            sheet_rels << { id: picture_rid, type: "#{DOC_REL}/image", target: "../media/sheet#{idx + 1}_bg.#{bg_ext}" }
            sheet[:_bg_media] = { path: "xl/media/sheet#{idx + 1}_bg.#{bg_ext}", data: bg_data.b, ext: bg_ext }
          end
        end

        sheet[:_prepared_rels] = sheet_rels
        sheet[:_enriched_hyperlinks] = enriched_hyperlinks

        {
          drawing_rid: drawing_rid,
          vml_rid: vml_rid,
          table_start_rid: table_start_rid,
          picture_rid: picture_rid,
          enriched_hyperlinks: enriched_hyperlinks,
          sheet_rels: sheet_rels
        }
      end

      def write_auxiliary_parts(zip)
        drawing_idx = 0
        comment_idx = 0
        tc_idx = 0
        table_idx = 0
        pivot_cache_idx = 0
        pivot_table_idx = 0

        @sheets.each_with_index do |sheet, idx|
          sheet_images = sheet[:images] || []
          sheet_charts = sheet[:charts] || []
          sheet_shapes = sheet[:shapes] || []
          sheet_comments = sheet[:comments] || []
          sheet_tables = sheet[:tables] || []
          has_drawing = sheet_images.any? || sheet_charts.any? || sheet_shapes.any?
          has_comments = sheet_comments.any?
          sheet_rels = sheet[:_prepared_rels] || []

          if has_drawing
            drawing_idx += 1
            drawing_rels_data = []
            drawing_parts = []

            chart_writer = Xlsxrb::Ooxml::Writer.new
            chart_writer.add_sheet(sheet[:name]) unless sheet[:name] == "Sheet1"

            (sheet[:rows] || []).each do |row|
              cells = row.is_a?(Hash) ? row[:cells] : row.cells
              (cells || []).each do |cell|
                ref = cell.is_a?(Hash) ? cell[:ref] : cell.ref
                value = cell.is_a?(Hash) ? cell[:value] : cell.value
                chart_writer.set_cell(ref, value, sheet: sheet[:name])
              end
            end
            sheet[:cells]&.each do |ref, value|
              chart_writer.set_cell(ref, value, sheet: sheet[:name])
            end

            sheet_images.each do |img|
              media_path = "xl/media/image#{drawing_idx}.#{img[:ext] || "png"}"
              zip.add_binary_entry(media_path, img[:file_data])
              drawing_rels_data << { type: :image, target: "../media/image#{drawing_idx}.#{img[:ext] || "png"}" }
              drawing_parts << { kind: :pic, img: img, rid_index: drawing_rels_data.size }
            end

            sheet_charts.each do |chart_options|
              chart_writer.add_chart(**chart_options)
            end

            processed_charts = chart_writer.charts
            processed_charts.each do |chart|
              chart_path = "xl/charts/chart#{drawing_rels_data.size + 1}.xml" # simplified
              zip.add_entry(chart_path, chart_writer.send(:generate_chart_xml, chart))
              drawing_rels_data << { type: :chart, target: "../charts/chart#{drawing_rels_data.size + 1}.xml" }
              drawing_parts << { kind: :chart, chart: chart, rid_index: drawing_rels_data.size }
            end

            sheet_shapes.each_with_index do |shape, si|
              drawing_parts << { kind: :sp, shape: shape, id: drawing_parts.size + si + 2 }
            end

            drawing_xml = chart_writer.send(:generate_drawing_xml, drawing_parts)
            zip.add_entry("xl/drawings/drawing#{drawing_idx}.xml", drawing_xml)
            unless drawing_rels_data.empty?
              drawing_rels_xml = chart_writer.send(:generate_drawing_rels, drawing_rels_data)
              zip.add_entry("xl/drawings/_rels/drawing#{drawing_idx}.xml.rels", drawing_rels_xml)
            end
          end

          if has_comments
            comment_idx += 1
            comment_writer = Xlsxrb::Ooxml::Writer.new
            sheet_comments.each do |c|
              c = c.to_h if c.is_a?(Elements::MetadataRecord)
              comment_writer.add_comment(c[:cell] || c[:ref], c[:text], author: c[:author] || "Author")
            end
            zip.add_entry("xl/comments#{comment_idx}.xml", comment_writer.send(:generate_comments_xml, comment_writer.comments))
            zip.add_entry("xl/drawings/vmlDrawing#{comment_idx}.vml", comment_writer.send(:generate_vml_drawing_xml, comment_writer.comments))
          end

          sheet_threaded_comments = sheet[:threaded_comments] || []
          if sheet_threaded_comments.any?
            tc_idx += 1
            tc_writer = Xlsxrb::Ooxml::Writer.new
            sheet_threaded_comments.each do |c|
              c = c.to_h if c.is_a?(Elements::MetadataRecord)
              tc_writer.add_threaded_comment(
                c[:cell] || c[:ref],
                c[:text],
                author: c[:author] || "Author",
                timestamp: c[:timestamp] || c[:dt],
                id: c[:id],
                parent_id: c[:parent_id],
                done: c[:done],
                user_id: c[:user_id],
                provider_id: c[:provider_id]
              )
            end
            tc_writer.persons.each do |p|
              @persons << p unless @persons.any? { |ep| ep[:id] == p[:id] || ep[:display_name] == p[:display_name] }
            end
            zip.add_entry("xl/threadedComments/threadedComment#{tc_idx}.xml", tc_writer.send(:generate_threaded_comments_xml, tc_writer.threaded_comments))
          end

          sheet_tables.each do |tbl|
            table_idx += 1
            zip.add_entry("xl/tables/table#{table_idx}.xml", build_table_xml(tbl, table_idx))
          end

          sheet_pivot_tables = sheet[:pivot_tables] || []
          unless sheet_pivot_tables.empty?
            pivot_writer = Xlsxrb::Ooxml::Writer.new
            pivot_writer.add_sheet(sheet[:name])
            sheet_pivot_tables.each do |pt|
              pivot_writer.add_pivot_table(
                pt[:source_ref],
                row_fields: pt[:row_fields],
                data_fields: pt[:data_fields],
                column_fields: pt[:column_fields] || pt[:col_fields] || [],
                destination: pt[:destination] || pt[:dest_ref] || pt[:ref] || "E1",
                name: pt[:name],
                field_names: pt[:field_names],
                items: pt[:items],
                sheet: sheet[:name]
              )
            end

            pivot_data = pivot_writer.pivot_tables(sheet: sheet[:name])
            pivot_data.each do |pt_data|
              pivot_cache_idx += 1
              pivot_table_idx += 1
              zip.add_entry("xl/pivotCache/pivotCacheDefinition#{pivot_cache_idx}.xml",
                            pivot_writer.send(:generate_pivot_cache_definition_xml, pt_data, pivot_cache_idx))
              zip.add_entry("xl/pivotCache/pivotCacheRecords#{pivot_cache_idx}.xml",
                            pivot_writer.send(:generate_pivot_cache_records_xml, pt_data))
              zip.add_entry("xl/pivotTables/pivotTable#{pivot_table_idx}.xml",
                            pivot_writer.send(:generate_pivot_table_xml, pt_data, pivot_cache_idx))
              zip.add_entry("xl/pivotCache/_rels/pivotCacheDefinition#{pivot_cache_idx}.xml.rels",
                            pivot_writer.send(:generate_pivot_cache_rels, pivot_cache_idx))
              zip.add_entry("xl/pivotTables/_rels/pivotTable#{pivot_table_idx}.xml.rels",
                            pivot_writer.send(:generate_pivot_table_rels, pivot_cache_idx))
            end
          end

          zip.add_binary_entry(sheet[:_bg_media][:path], sheet[:_bg_media][:data]) if sheet[:_bg_media]

          zip.add_entry("xl/worksheets/_rels/sheet#{idx + 1}.xml.rels", build_sheet_rels_from_list(sheet_rels)) unless sheet_rels.empty?
        end

        return unless @persons.any?

        p_writer = Xlsxrb::Ooxml::Writer.new
        zip.add_entry("xl/persons/person.xml", p_writer.send(:generate_persons_xml, @persons))
      end

      def write_package_metadata(zip)
        zip.add_entry("[Content_Types].xml", build_content_types)
        zip.add_entry("_rels/.rels", build_root_rels)
        zip.add_entry("xl/workbook.xml", build_workbook_xml)
        zip.add_entry("xl/_rels/workbook.xml.rels", build_workbook_rels)
        zip.add_entry("xl/styles.xml", build_styles_xml)
        write_shared_strings_xml(zip) unless @shared_strings.empty?

        # Document properties
        zip.add_entry("docProps/core.xml", build_core_properties_xml) unless @core_properties.empty?
        zip.add_entry("docProps/app.xml", build_app_properties_xml) unless @app_properties.empty?
        zip.add_entry("docProps/custom.xml", build_custom_properties_xml) unless @custom_properties.empty?
      end

      def write_package_parts(zip)
        preprocess_conditional_formats!
        write_auxiliary_parts(zip)
        write_package_metadata(zip)
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
        pivot_cache_count = 0
        pivot_table_count = 0
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
              next if image_exts[ext]

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

          if sheet[:_bg_media]
            bg_ext = sheet[:_bg_media][:ext]
            unless image_exts[bg_ext]
              mime = case bg_ext
                     when "png" then "image/png"
                     when "jpg", "jpeg" then "image/jpeg"
                     when "gif" then "image/gif"
                     when "bmp" then "image/bmp"
                     when "emf" then "image/x-emf"
                     when "wmf" then "image/x-wmf"
                     else "application/octet-stream"
                     end
              b.empty_tag("Default", { Extension: bg_ext, ContentType: mime })
              image_exts[bg_ext] = true
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

          sheet_pivots = sheet[:pivot_tables] || []
          sheet_pivots.each do |_pt|
            pivot_cache_count += 1
            pivot_table_count += 1
            b.empty_tag("Override", { PartName: "/xl/pivotCache/pivotCacheDefinition#{pivot_cache_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml" })
            b.empty_tag("Override", { PartName: "/xl/pivotCache/pivotCacheRecords#{pivot_cache_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml" })
            b.empty_tag("Override", { PartName: "/xl/pivotTables/pivotTable#{pivot_table_count}.xml", ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml" })
          end
        end

        b.empty_tag("Override", { PartName: "/xl/persons/person.xml", ContentType: "application/vnd.ms-excel.person+xml" }) if @persons.any?
        (1..@threaded_comment_count).each do |tc|
          b.empty_tag("Override", { PartName: "/xl/threadedComments/threadedComment#{tc}.xml", ContentType: "application/vnd.ms-excel.threadedcomments+xml" })
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

        # Workbook properties
        if @workbook_properties && !@workbook_properties.empty?
          attrs = {}
          attrs[:date1904] = "1" if @workbook_properties[:date1904]
          attrs[:updateLinks] = @workbook_properties[:update_links] if @workbook_properties[:update_links]
          attrs[:codeName] = @workbook_properties[:code_name] if @workbook_properties[:code_name]
          b.empty_tag("workbookPr", attrs) unless attrs.empty?
        end

        # Workbook protection
        if @workbook_protection
          wp_attrs = {}
          wp_attrs[:lockStructure] = "1" if @workbook_protection[:lock_structure]
          wp_attrs[:lockWindows] = "1" if @workbook_protection[:lock_windows]
          wp_attrs[:workbookPassword] = @workbook_protection[:password] if @workbook_protection[:password]
          b.empty_tag("workbookProtection", wp_attrs) unless wp_attrs.empty?
        end

        # Workbook view: activeTab, firstSheet, showSheetTabs, tabRatio
        active_tab_idx = resolve_active_tab_index
        first_sheet_idx = resolve_first_sheet_index
        show_sheet_tabs = @workbook_properties&.[](:show_sheet_tabs)
        tab_ratio = @workbook_properties&.[](:tab_ratio)

        emit_active_tab = !(@workbook_properties&.[](:active_sheet_index) || @workbook_properties&.[](:active_sheet) || @workbook_properties&.[](:active_tab)).nil? || active_tab_idx&.positive?

        has_book_views = emit_active_tab || first_sheet_idx || !show_sheet_tabs.nil? || tab_ratio
        if has_book_views
          bv_attrs = {}
          bv_attrs[:activeTab] = active_tab_idx.to_s if emit_active_tab
          bv_attrs[:firstSheet] = first_sheet_idx.to_s if first_sheet_idx
          bv_attrs[:showSheetTabs] = show_sheet_tabs ? "1" : "0" unless show_sheet_tabs.nil?
          bv_attrs[:tabRatio] = tab_ratio.to_s if tab_ratio
          b.open_tag("bookViews")
          b.empty_tag("workbookView", bv_attrs)
          b.close_tag("bookViews")
        end

        b.open_tag("sheets")
        @sheets.each_with_index do |sheet, idx|
          s_name = sheet.respond_to?(:name) ? sheet.name : sheet[:name]
          sheet_attrs = {
            name: s_name,
            sheetId: (idx + 1).to_s,
            "r:id": "rId#{idx + 1}"
          }
          state = sheet.respond_to?(:state) ? sheet.state : sheet[:state]
          if state && state.to_sym != :visible
            sheet_attrs[:state] = state.to_sym == :very_hidden ? "veryHidden" : "hidden"
          end
          b.empty_tag("sheet", sheet_attrs)
        end
        b.close_tag("sheets")

        # Defined names
        unless @defined_names.empty?
          b.open_tag("definedNames")
          @defined_names.each do |dn|
            dn = dn.to_h if dn.is_a?(Elements::MetadataRecord)
            dn_attrs = { name: dn[:name] }
            local_id = dn[:local_sheet_id]
            target_name = (dn[:sheet] || dn[:local_sheet_name])&.to_s
            if !local_id && target_name
              matched_idx = @sheets.find_index { |s| (s.respond_to?(:name) ? s.name : s[:name]) == target_name }
              local_id = matched_idx if matched_idx
            end
            dn_attrs[:localSheetId] = local_id.to_s if local_id
            dn_attrs[:hidden] = "1" if dn[:hidden]
            val = dn[:value] || (target_name && dn[:range] ? "'#{target_name}'!#{dn[:range]}" : dn[:range])
            b.tag("definedName", dn_attrs) { |_| b.text(val) }
          end
          b.close_tag("definedNames")
        end

        # Calculation properties: ask Excel to recalculate every formula when the file is opened
        b.empty_tag("calcPr", { fullCalcOnLoad: "1" }) if @workbook_properties&.[](:full_calc_on_load)

        # Pivot caches
        total_pivots = @sheets.sum { |s| (s[:pivot_tables] || []).size }
        if total_pivots.positive?
          # rId for pivot caches starts after sheets + styles + sharedStrings
          # In build_workbook_rels: sheets(1..N), styles(N+1), sharedStrings(N+2 if present)
          pivot_rid_base = @sheets.size + 1 + (@shared_strings.empty? ? 0 : 1)
          b.open_tag("pivotCaches")
          total_pivots.times do |i|
            b.empty_tag("pivotCache", { cacheId: (i + 1).to_s, "r:id": "rId#{pivot_rid_base + i + 1}" })
          end
          b.close_tag("pivotCaches")
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

        # Pivot cache relationships
        total_pivots = @sheets.sum { |s| (s[:pivot_tables] || []).size }
        total_pivots.times do |i|
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "#{DOC_REL}/pivotCacheDefinition",
                        Target: "pivotCache/pivotCacheDefinition#{i + 1}.xml"
                      })
        end

        if @persons.any?
          rid += 1
          b.empty_tag("Relationship", {
                        Id: "rId#{rid}",
                        Type: "http://schemas.microsoft.com/office/2017/10/relationships/person",
                        Target: "persons/person.xml"
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

        # Number formats
        raw_num_fmts = if @styles.respond_to?(:number_formats)
                         @styles.number_formats
                       else
                         @styles&.dig(:num_fmts) || []
                       end
        num_fmts = if raw_num_fmts.is_a?(Hash)
                     raw_num_fmts.map { |k, v| { num_fmt_id: k.to_i, format_code: v.to_s } }
                   elsif raw_num_fmts.is_a?(Array)
                     raw_num_fmts.map do |nf|
                       if nf.is_a?(Hash)
                         id = (nf[:num_fmt_id] || nf[:id]).to_i
                         code = (nf[:format_code] || nf[:code] || nf[:format]).to_s
                         { num_fmt_id: id, format_code: code }
                       elsif nf.is_a?(Array) && nf.size == 2
                         { num_fmt_id: nf[0].to_i, format_code: nf[1].to_s }
                       end
                     end.compact
                   else
                     []
                   end
        unless num_fmts.empty?
          b.tag("numFmts", { count: num_fmts.size.to_s }) do |_|
            num_fmts.each do |nf|
              b.empty_tag("numFmt", { numFmtId: nf[:num_fmt_id].to_s, formatCode: nf[:format_code] })
            end
          end
        end

        # Fonts
        fonts = if @styles.respond_to?(:fonts)
                  @styles.fonts
                else
                  @styles&.dig(:fonts) || []
                end
        fonts = [{}] if fonts.empty? # At least one default font
        b.tag("fonts", { count: fonts.size.to_s }) do |_|
          fonts.each do |font_props|
            font_props = font_props.to_h if font_props.is_a?(Elements::MetadataRecord)
            b.tag("font") do |_|
              b.empty_tag("b") if font_props[:bold]
              b.empty_tag("i") if font_props[:italic]
              b.empty_tag("strike") if font_props[:strike]
              if font_props[:underline]
                u_val = font_props[:underline]
                if u_val == true || u_val.to_s == "single"
                  b.empty_tag("u")
                else
                  mapped_u = StyleBuilder::UNDERLINE_MAP[u_val.to_sym] || u_val.to_s
                  b.empty_tag("u", { val: mapped_u })
                end
              end
              va_val = font_props[:vert_align] || font_props[:vertical_align]
              va_val ||= "superscript" if font_props[:superscript]
              va_val ||= "subscript" if font_props[:subscript]
              if va_val
                mapped_va = StyleBuilder::VERT_ALIGN_MAP[va_val.to_sym] || va_val.to_s
                b.empty_tag("vertAlign", { val: mapped_va })
              end
              sz_val = font_props[:size] || font_props[:sz]
              b.empty_tag("sz", { val: (sz_val || 11).to_s }) if sz_val || fonts == [{}]
              color_attrs = color_tag_attrs(font_props[:color] || (font_props.key?(:rgb) || font_props.key?(:theme) || font_props.key?(:indexed) ? font_props : nil))
              b.empty_tag("color", color_attrs) if color_attrs
              b.empty_tag("name", { val: font_props[:name] || "Calibri" })
              b.empty_tag("family", { val: font_props[:family].to_s }) if font_props[:family]
              b.empty_tag("charset", { val: font_props[:charset].to_s }) if font_props[:charset]
              b.empty_tag("scheme", { val: font_props[:scheme].to_s }) if font_props[:scheme]
            end
          end
        end

        # Fills
        fills = if @styles.respond_to?(:fills)
                  @styles.fills
                else
                  @styles&.dig(:fills) || []
                end
        fills = [{ pattern: "none" }, { pattern: "gray125" }] if fills.empty?
        b.tag("fills", { count: fills.size.to_s }) do |_|
          fills.each do |fill_props|
            fill_props = fill_props.to_h if fill_props.is_a?(Elements::MetadataRecord)
            b.tag("fill") do |_|
              if fill_props[:gradient]
                b.open_tag("gradientFill", { type: fill_props[:gradient][:type], degree: fill_props[:gradient][:degree]&.to_s }.compact)
                fill_props[:gradient][:stops].each do |stop|
                  b.tag("stop", { position: stop[:position].to_s }) do |_|
                    c_attrs = color_tag_attrs(stop[:color] || stop)
                    b.empty_tag("color", c_attrs) if c_attrs
                  end
                end
                b.close_tag("gradientFill")
              else
                b.tag("patternFill", { patternType: fill_props[:pattern] || "none" }) do |_|
                  fg_attrs = color_tag_attrs(fill_props[:fg_color] || (fill_props[:fg_color_theme] ? { theme: fill_props[:fg_color_theme], tint: fill_props[:fg_color_tint] } : nil) || (fill_props[:fg_color_indexed] ? { indexed: fill_props[:fg_color_indexed] } : nil))
                  b.empty_tag("fgColor", fg_attrs) if fg_attrs

                  bg_attrs = color_tag_attrs(fill_props[:bg_color] || (fill_props[:bg_color_theme] ? { theme: fill_props[:bg_color_theme], tint: fill_props[:bg_color_tint] } : nil) || (fill_props[:bg_color_indexed] ? { indexed: fill_props[:bg_color_indexed] } : nil))
                  b.empty_tag("bgColor", bg_attrs) if bg_attrs
                end
              end
            end
          end
        end

        # Borders
        borders = if @styles.respond_to?(:borders)
                    @styles.borders
                  else
                    @styles&.dig(:borders) || []
                  end
        borders = [{}] if borders.empty?
        b.tag("borders", { count: borders.size.to_s }) do |_|
          borders.each do |border_props|
            border_props = border_props.to_h if border_props.is_a?(Elements::MetadataRecord)
            border_attrs = {}
            border_attrs[:diagonalUp] = "1" if border_props[:diagonal_up]
            border_attrs[:diagonalDown] = "1" if border_props[:diagonal_down]
            b.tag("border", border_attrs) do |_|
              %i[left right top bottom diagonal].each do |side|
                side_data = border_props[side]
                if side_data
                  b.tag(side.to_s, { style: side_data[:style] }) do |_|
                    c_attrs = color_tag_attrs(side_data[:color])
                    b.empty_tag("color", c_attrs) if c_attrs
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

        # Cell XFs (formatting)
        xf_entries = if @styles.respond_to?(:cell_xfs)
                       @styles.cell_xfs
                     else
                       @styles&.dig(:cell_xfs) || @styles&.dig(:xf_entries) || []
                     end
        xf_count = [xf_entries.size, 1].max
        b.tag("cellXfs", { count: xf_count.to_s }) do |_|
          xf_entries.each do |xf|
            xf = xf.to_h if xf.is_a?(Elements::MetadataRecord)
            attrs = {
              numFmtId: (xf[:num_fmt_id] || 0).to_s,
              fontId: (xf[:font_id] || 0).to_s,
              fillId: (xf[:fill_id] || 0).to_s,
              borderId: (xf[:border_id] || 0).to_s
            }
            attrs[:xfId] = xf[:xf_id].to_s if xf[:xf_id]
            attrs[:applyNumberFormat] = "1" if xf[:num_fmt_id]&.to_i&.positive?
            attrs[:applyFont] = "1" if xf[:font_id]&.to_i&.positive?
            attrs[:applyFill] = "1" if xf[:fill_id]&.to_i&.positive?
            attrs[:applyBorder] = "1" if xf[:border_id]&.to_i&.positive?
            attrs[:applyAlignment] = "1" if xf[:alignment]
            attrs[:applyProtection] = "1" if xf[:protection]

            if xf[:alignment] || xf[:protection]
              b.tag("xf", attrs) do |_|
                if xf[:alignment]
                  align_attrs = {}
                  align_attrs[:horizontal] = xf[:alignment][:horizontal] if xf[:alignment][:horizontal]
                  align_attrs[:vertical] = xf[:alignment][:vertical] if xf[:alignment][:vertical]
                  align_attrs[:wrapText] = "1" if xf[:alignment][:wrap_text]
                  align_attrs[:textRotation] = xf[:alignment][:text_rotation].to_s if xf[:alignment][:text_rotation]
                  align_attrs[:indent] = xf[:alignment][:indent].to_s if xf[:alignment][:indent]
                  align_attrs[:relativeIndent] = xf[:alignment][:relative_indent].to_s if xf[:alignment][:relative_indent]
                  align_attrs[:shrinkToFit] = "1" if xf[:alignment][:shrink_to_fit]
                  align_attrs[:readingOrder] = xf[:alignment][:reading_order].to_s if xf[:alignment][:reading_order]
                  align_attrs[:justifyLastLine] = "1" if xf[:alignment][:justify_last_line]
                  b.empty_tag("alignment", align_attrs)
                end
                if xf[:protection]
                  prot_attrs = {}
                  prot_attrs[:locked] = xf[:protection][:locked] ? "1" : "0" unless xf[:protection][:locked].nil?
                  prot_attrs[:hidden] = xf[:protection][:hidden] ? "1" : "0" unless xf[:protection][:hidden].nil?
                  b.empty_tag("protection", prot_attrs)
                end
              end
            else
              b.empty_tag("xf", attrs)
            end
          end
          # Add default if empty
          b.empty_tag("xf", { numFmtId: "0", fontId: "0", fillId: "0", borderId: "0" }) if xf_entries.empty?
        end

        # Differential formats (for conditional formatting)
        dxfs = if @styles.respond_to?(:dxfs)
                 @styles.dxfs
               else
                 @styles&.dig(:dxfs) || []
               end
        unless dxfs.empty?
          b.tag("dxfs", { count: dxfs.size.to_s }) do |_|
            dxfs.each do |dxf|
              d = dxf.is_a?(Elements::MetadataRecord) ? dxf.to_h : dxf
              b.tag("dxf") do |_|
                write_dxf_font(b, d[:font]) if d[:font]
                nf = d[:number_format] || d[:num_fmt]
                if nf
                  nf_h = nf.is_a?(Elements::MetadataRecord) ? nf.to_h : nf
                  b.empty_tag("numFmt", { numFmtId: nf_h[:num_fmt_id].to_s, formatCode: nf_h[:format_code].to_s })
                end
                write_dxf_fill(b, d[:fill]) if d[:fill]
                write_dxf_alignment(b, d[:alignment]) if d[:alignment]
                write_dxf_border(b, d[:border]) if d[:border]
                write_dxf_protection(b, d[:protection]) if d[:protection]
              end
            end
          end
        end

        b.close_tag("styleSheet")
        io.string
      end

      def preprocess_conditional_formats!
        @sheets.each do |sheet|
          rules = sheet[:conditional_formats]
          next if rules.nil? || rules.empty?

          sheet[:conditional_formats] = rules.map do |rule|
            next rule if rule.respond_to?(:dxf_id) ? rule.dxf_id : (rule.is_a?(Hash) && (rule[:dxf_id] || rule[:format_id]))

            normalized = rule.is_a?(Hash) ? rule.dup : rule.to_h
            dxf = normalized.delete(:dxf) || extract_dxf_from_rule!(normalized)
            next normalized unless dxf

            dxf_id = if @styles.respond_to?(:add_dxf)
                       @styles.dxfs.index(dxf) || @styles.add_dxf(dxf)
                     else
                       @styles[:dxfs] ||= []
                       @styles[:dxfs].index(dxf) || begin
                         @styles[:dxfs] << dxf
                         @styles[:dxfs].size - 1
                       end
                     end
            normalized[:dxf_id] = dxf_id
            normalized
          end
        end
      end

      def extract_dxf_from_rule!(rule)
        DslHelpers.extract_dxf_from_rule!(rule)
      end

      def write_dxf_font(builder, font)
        font = font.to_h if font.is_a?(Elements::MetadataRecord)
        builder.tag("font") do |_|
          if font.key?(:bold)
            font[:bold] ? builder.empty_tag("b") : builder.empty_tag("b", { val: "0" })
          end
          if font.key?(:italic)
            font[:italic] ? builder.empty_tag("i") : builder.empty_tag("i", { val: "0" })
          end
          if font.key?(:strike)
            font[:strike] ? builder.empty_tag("strike") : builder.empty_tag("strike", { val: "0" })
          end
          if font[:underline]
            if [true, "single"].include?(font[:underline])
              builder.empty_tag("u")
            else
              mapped_u = StyleBuilder::UNDERLINE_MAP[font[:underline].to_sym] || font[:underline]
              builder.empty_tag("u", { val: mapped_u })
            end
          end
          va_val = font[:vert_align] || font[:vertical_align]
          va_val ||= "superscript" if font[:superscript]
          va_val ||= "subscript" if font[:subscript]
          if va_val
            mapped_va = StyleBuilder::VERT_ALIGN_MAP[va_val.to_sym] || va_val.to_s
            builder.empty_tag("vertAlign", { val: mapped_va })
          end
          color_attrs = color_tag_attrs(font[:color] || font)
          builder.empty_tag("color", color_attrs) if color_attrs
          builder.empty_tag("name", { val: font[:name] }) if font[:name]
          sz_val = font[:sz] || font[:size]
          builder.empty_tag("sz", { val: sz_val.to_s }) if sz_val
        end
      end

      def write_dxf_fill(builder, fill)
        fill = fill.to_h if fill.is_a?(Elements::MetadataRecord)
        builder.tag("fill") do |_|
          pattern_attrs = {}
          pattern = fill[:pattern] || fill[:pattern_type]
          pattern ||= "solid" if (fill[:color] || fill[:fill_color] || fill[:fg_color]) && !fill[:bg_color]
          pattern_attrs[:patternType] = pattern if pattern
          builder.tag("patternFill", pattern_attrs) do |_|
            fg = fill[:fg_color] || fill[:color] || fill[:fill_color]
            fg_attrs = color_tag_attrs(fg)
            builder.empty_tag("fgColor", fg_attrs) if fg_attrs
            bg_attrs = color_tag_attrs(fill[:bg_color])
            builder.empty_tag("bgColor", bg_attrs) if bg_attrs
          end
        end
      end

      def write_dxf_alignment(builder, alignment)
        alignment = alignment.to_h if alignment.is_a?(Elements::MetadataRecord)
        return if alignment.nil? || alignment.empty?

        align_attrs = {}
        align_attrs[:horizontal] = alignment[:horizontal] if alignment[:horizontal]
        align_attrs[:vertical] = alignment[:vertical] if alignment[:vertical]
        align_attrs[:wrapText] = "1" if alignment[:wrap_text]
        align_attrs[:textRotation] = alignment[:text_rotation].to_s if alignment[:text_rotation]
        align_attrs[:indent] = alignment[:indent].to_s if alignment[:indent]
        align_attrs[:relativeIndent] = alignment[:relative_indent].to_s if alignment[:relative_indent]
        align_attrs[:shrinkToFit] = "1" if alignment[:shrink_to_fit]
        align_attrs[:readingOrder] = alignment[:reading_order].to_s if alignment[:reading_order]
        align_attrs[:justifyLastLine] = "1" if alignment[:justify_last_line]
        builder.empty_tag("alignment", align_attrs) unless align_attrs.empty?
      end

      def write_dxf_border(builder, border)
        border = border.to_h if border.is_a?(Elements::MetadataRecord)
        return if border.nil? || border.empty?

        border_attrs = {}
        border_attrs[:diagonalUp] = "1" if border[:diagonal_up]
        border_attrs[:diagonalDown] = "1" if border[:diagonal_down]

        builder.tag("border", border_attrs) do |_|
          %i[left right top bottom diagonal].each do |side|
            side_data = border[side]
            next unless side_data

            tag_attrs = {}
            tag_attrs[:style] = side_data[:style] if side_data[:style]
            c_attrs = color_tag_attrs(side_data[:color])
            if c_attrs
              builder.tag(side.to_s, tag_attrs) do |_|
                builder.empty_tag("color", c_attrs)
              end
            elsif !tag_attrs.empty?
              builder.empty_tag(side.to_s, tag_attrs)
            else
              builder.empty_tag(side.to_s)
            end
          end
        end
      end

      def write_dxf_protection(builder, protection)
        protection = protection.to_h if protection.is_a?(Elements::MetadataRecord)
        return if protection.nil? || protection.empty?

        prot_attrs = {}
        prot_attrs[:locked] = protection[:locked] ? "1" : "0" unless protection[:locked].nil?
        prot_attrs[:hidden] = protection[:hidden] ? "1" : "0" unless protection[:hidden].nil?
        builder.empty_tag("protection", prot_attrs) unless prot_attrs.empty?
      end

      # : (untyped rgb) -> ::String?
      def format_rgb_hex(rgb)
        return nil unless rgb

        str = rgb.to_s.sub(/^#/, "").strip
        return nil if str.empty?

        case str
        when /\A[0-9a-fA-F]{8}\z/
          str.upcase
        when /\A[0-9a-fA-F]{6}\z/
          "FF#{str.upcase}"
        when /\A[0-9a-fA-F]{3}\z/
          r, g, b = str.chars
          "FF#{r}#{r}#{g}#{g}#{b}#{b}".upcase
        else
          resolved = StyleBuilder.resolve_color(rgb)
          case resolved
          when /\A[0-9a-fA-F]{8}\z/
            resolved.upcase
          when /\A[0-9a-fA-F]{6}\z/
            "FF#{resolved.upcase}"
          end
        end
      end

      # : (untyped spec) -> ::Hash[Symbol, ::String]?
      def color_tag_attrs(spec)
        return nil unless spec

        spec = spec[:color] if spec.is_a?(Hash) && spec[:color].is_a?(Hash)
        return nil unless spec

        if spec.is_a?(Hash)
          if spec[:auto]
            { auto: "1" }
          elsif spec[:rgb] || (spec[:color] && !spec[:color].is_a?(Hash))
            raw_rgb = spec[:rgb] || spec[:color]
            fmt = format_rgb_hex(raw_rgb)
            fmt ? { rgb: fmt } : nil
          elsif spec[:theme]
            attrs = { theme: spec[:theme].to_s }
            attrs[:tint] = spec[:tint].to_s if spec[:tint]
            attrs
          elsif spec[:indexed]
            { indexed: spec[:indexed].to_s }
          end
        elsif spec.is_a?(String) || spec.is_a?(Symbol)
          fmt = format_rgb_hex(spec)
          fmt ? { rgb: fmt } : nil
        end
      end

      INVALID_XML_CHARS_RE = /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/
      ESCAPE_RE = /[&<>"']/
      ESCAPE_MAP = { "&" => "&amp;", "<" => "&lt;", ">" => "&gt;", '"' => "&quot;", "'" => "&apos;" }.freeze
      private_constant :INVALID_XML_CHARS_RE, :ESCAPE_RE, :ESCAPE_MAP

      def escape_xml(value)
        str = value.to_s
        str = str.dup.force_encoding(Encoding::UTF_8) if str.encoding == Encoding::BINARY
        str = str.gsub(INVALID_XML_CHARS_RE, "") if str.match?(INVALID_XML_CHARS_RE)
        str.match?(ESCAPE_RE) ? str.gsub(ESCAPE_RE, ESCAPE_MAP) : str
      end

      def write_shared_strings_xml(zip)
        return if @shared_strings.empty?

        zip.start_entry("xl/sharedStrings.xml")
        count_str = @shared_strings.size.to_s
        buf = String.new(capacity: 65_536)
        buf << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        buf << '<sst xmlns="' << SSML_NS << '" count="' << count_str << '" uniqueCount="' << count_str << '">'

        @shared_strings.each do |str|
          if str.is_a?(Xlsxrb::Elements::RichText)
            buf << "<si>"
            str.runs.each do |run|
              font = run.respond_to?(:font) ? run.font : run[:font]
              text = run.respond_to?(:text) ? run.text : run[:text]
              if font && !font.empty?
                buf << "<r><rPr>"
                buf << "<b/>" if font[:bold]
                buf << "<i/>" if font[:italic]
                buf << "<strike/>" if font[:strike]
                if font[:underline]
                  if font[:underline] == true || font[:underline].to_s == "single"
                    buf << "<u/>"
                  else
                    mapped_u = StyleBuilder::UNDERLINE_MAP[font[:underline].to_sym] || font[:underline]
                    buf << '<u val="' << mapped_u.to_s << '"/>'
                  end
                end
                va_val = font[:vert_align] || font[:vertical_align]
                va_val ||= "superscript" if font[:superscript]
                va_val ||= "subscript" if font[:subscript]
                if va_val
                  mapped_va = StyleBuilder::VERT_ALIGN_MAP[va_val.to_sym] || va_val.to_s
                  buf << '<vertAlign val="' << mapped_va << '"/>'
                end
                font_sz = font[:sz] || font[:size]
                buf << '<sz val="' << font_sz.to_s << '"/>' if font_sz
                color_attrs = color_tag_attrs(font[:color] || font)
                if color_attrs
                  attrs_str = color_attrs.map { |k, v| %(#{k}="#{v}") }.join(" ")
                  buf << "<color #{attrs_str}/>"
                end
                buf << '<rFont val="' << escape_xml(font[:name]) << '"/>' if font[:name]
                buf << '<family val="' << font[:family].to_s << '"/>' if font[:family]
                buf << '<scheme val="' << font[:scheme].to_s << '"/>' if font[:scheme]
                buf << "</rPr><t>"
              else
                buf << "<r><t>"
              end
              buf << escape_xml(text.to_s) << "</t></r>"
            end
            buf << "</si>"
          else
            buf << "<si><t>" << escape_xml(str.to_s) << "</t></si>"
          end

          if buf.bytesize >= 32_768
            zip.write_data(buf)
            buf.clear
          end
        end

        buf << "</sst>"
        zip.write_data(buf)
        buf.clear
        zip.finish_entry
      end

      def write_worksheet_xml(io, sheet, drawing_rid: nil, picture_rid: nil, sheet_protection: nil,
                              protected_ranges: nil, auto_filter: nil, filter_columns: nil,
                              sort_state: nil, merge_cells: nil, conditional_formats: nil,
                              data_validations: nil, hyperlinks: nil,
                              print_options: nil, page_margins: nil, page_setup: nil,
                              header_footer: nil, row_breaks: nil, col_breaks: nil, column_breaks: nil,
                              cell_watches: nil, ignored_errors: nil, freeze_pane: nil, split_pane: nil,
                              selection: nil, sheet_view: nil, sheet_properties: nil,
                              sheet_format: nil, tables: nil, table_start_rid: nil,
                              legacy_drawing_rid: nil, sparkline_groups: nil, phonetic_pr: nil,
                              sheet_calc_pr: nil)
        date1904 = @workbook_properties&.[](:date1904) ? true : false
        ws = WorksheetWriter.new(io, date1904: date1904)
        ws.start(
          columns: sheet[:columns] || [],
          sheet_properties: sheet_properties,
          freeze_pane: freeze_pane,
          split_pane: split_pane,
          selection: selection,
          sheet_view: sheet_view,
          sheet_format: sheet_format
        )
        if sheet[:rows_tmp_path]
          File.open(sheet[:rows_tmp_path], "rb") do |f|
            while (chunk = f.read(16_384))
              io.write(chunk)
            end
          end
        else
          (sheet[:rows] || []).each do |row|
            if row.is_a?(Hash)
              ws.write_row(row[:index], row[:cells], attrs: row[:attrs] || {}, unmapped: row[:unmapped] || [], sst_index: @shared_strings_index)
            else
              attrs = {}
              attrs[:height] = row.height if row.height
              attrs[:hidden] = true if row.hidden?
              attrs[:custom_height] = true if row.custom_height?
              attrs[:outline_level] = row.outline_level if row.outline_level
              style_idx = row.unmapped_data&.[](:style_index) || (row.respond_to?(:style_index) && row.style_index)
              attrs[:style_index] = style_idx if style_idx
              ws.write_row(row.index, row.cells, attrs: attrs, unmapped: row.unmapped_data || [], sst_index: @shared_strings_index)
            end
          end
        end
        ws.finish(
          drawing_rid: drawing_rid,
          picture_rid: picture_rid,
          sheet_protection: sheet_protection,
          protected_ranges: protected_ranges,
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
          col_breaks: col_breaks || column_breaks,
          cell_watches: cell_watches,
          ignored_errors: ignored_errors,
          tables: tables,
          table_start_rid: table_start_rid,
          legacy_drawing_rid: legacy_drawing_rid,
          sparkline_groups: sparkline_groups,
          phonetic_pr: phonetic_pr,
          sheet_calc_pr: sheet_calc_pr
        )
      end

      def build_table_xml(tbl, table_id)
        tbl = tbl.to_h if tbl.is_a?(Elements::MetadataRecord)
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
        t_attrs[:totalsRowCount] = tbl[:totals_row_count].to_s if tbl[:totals_row_count]&.positive?
        b.open_tag("table", t_attrs)

        # Auto filter for the table range unless auto_filter: false
        b.empty_tag("autoFilter", { ref: tbl[:ref] }) unless tbl[:auto_filter] == false

        # Table columns
        columns = tbl[:columns] || []
        b.open_tag("tableColumns", { count: columns.size.to_s })
        columns.each_with_index do |col, ci|
          col = col.to_h if col.is_a?(Elements::MetadataRecord)
          col_name = if col.is_a?(Hash)
                       col[:name]
                     elsif col.is_a?(Elements::TableColumn)
                       col.name
                     else
                       col
                     end
          col_attrs = { id: (ci + 1).to_s, name: col_name.to_s }
          calc_formula = nil
          tot_formula = nil
          if col.is_a?(Hash) || col.is_a?(Elements::TableColumn)
            calc_formula = col[:calculated_column_formula] || col[:data_formula] || col[:formula]
            tot_formula = col[:totals_row_formula]
            trf = col[:totals_row_function] || ("custom" if tot_formula)
            col_attrs[:totalsRowFunction] = trf.to_s if trf
            col_attrs[:totalsRowLabel] = col[:totals_row_label].to_s if col[:totals_row_label]
            col_attrs[:dataDxfId] = col[:data_dxf_id].to_s if col[:data_dxf_id]
            col_attrs[:totalsRowDxfId] = col[:totals_row_dxf_id].to_s if col[:totals_row_dxf_id]
            col_attrs[:headerRowDxfId] = col[:header_row_dxf_id].to_s if col[:header_row_dxf_id]
            col_attrs[:dataCellStyle] = col[:data_cell_style].to_s if col[:data_cell_style]
          end
          if calc_formula || tot_formula
            b.open_tag("tableColumn", col_attrs)
            b.tag("calculatedColumnFormula") { |cb| cb.text(calc_formula.to_s) } if calc_formula
            b.tag("totalsRowFormula") { |tb| tb.text(tot_formula.to_s) } if tot_formula
            b.close_tag("tableColumn")
          else
            b.empty_tag("tableColumn", col_attrs)
          end
        end
        b.close_tag("tableColumns")

        # Table style. Accepts either:
        # - style: "TableStyleMedium9"
        # - style: { name:, show_* }
        # and also show_* values at the table root for Facade compatibility.
        raw_style = tbl[:style] || tbl[:style_info]
        raw_style = raw_style.to_h if raw_style.respond_to?(:to_h)
        style = if raw_style.is_a?(Hash)
                  raw_style
                elsif raw_style
                  { name: raw_style.to_s }
                else
                  {}
                end

        style_attrs = { name: style[:name] || "TableStyleMedium2" }

        show_first_column = style.key?(:show_first_column) ? style[:show_first_column] : tbl[:show_first_column]
        show_last_column = style.key?(:show_last_column) ? style[:show_last_column] : tbl[:show_last_column]
        show_row_stripes = if style.key?(:show_row_stripes)
                             style[:show_row_stripes]
                           elsif tbl.key?(:show_row_stripes)
                             tbl[:show_row_stripes]
                           else
                             true
                           end
        show_column_stripes = style.key?(:show_column_stripes) ? style[:show_column_stripes] : tbl[:show_column_stripes]

        style_attrs[:showFirstColumn] = show_first_column ? "1" : "0" unless show_first_column.nil?
        style_attrs[:showLastColumn] = show_last_column ? "1" : "0" unless show_last_column.nil?
        style_attrs[:showRowStripes] = show_row_stripes ? "1" : "0"
        style_attrs[:showColumnStripes] = show_column_stripes ? "1" : "0" unless show_column_stripes.nil?
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
        b.tag("cp:revision") { |_| b.text(@core_properties[:revision]) } if @core_properties[:revision]
        if @core_properties[:created]
          created_val = @core_properties[:created].respond_to?(:iso8601) ? @core_properties[:created].iso8601 : @core_properties[:created]
          b.tag("dcterms:created", { "xsi:type": "dcterms:W3CDTF" }) { |_| b.text(created_val) }
        end
        if @core_properties[:modified]
          modified_val = @core_properties[:modified].respond_to?(:iso8601) ? @core_properties[:modified].iso8601 : @core_properties[:modified]
          b.tag("dcterms:modified", { "xsi:type": "dcterms:W3CDTF" }) { |_| b.text(modified_val) }
        end
        b.tag("cp:category") { |_| b.text(@core_properties[:category]) } if @core_properties[:category]
        b.tag("cp:contentStatus") { |_| b.text(@core_properties[:content_status]) } if @core_properties[:content_status]
        b.tag("dc:language") { |_| b.text(@core_properties[:language]) } if @core_properties[:language]
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
        app_val = @app_properties[:application] || @app_properties[:app_name]
        b.tag("Application") { |_| b.text(app_val) } if app_val
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
        list = if @custom_properties.is_a?(Hash)
                 @custom_properties.map do |name, val|
                   if val.is_a?(Elements::CustomProperty) || (val.respond_to?(:name) && val.respond_to?(:value))
                     { name: val.name || name.to_s, value: val.value, type: val.type }
                   elsif val.is_a?(Hash)
                     { name: val[:name] || name.to_s, value: val[:value], type: val[:type] }
                   else
                     { name: name.to_s, value: val, type: nil }
                   end
                 end
               else
                 @custom_properties.map do |prop|
                   if prop.is_a?(Elements::CustomProperty) || (prop.respond_to?(:name) && prop.respond_to?(:value))
                     { name: prop.name, value: prop.value, type: prop.type }
                   else
                     prop
                   end
                 end
               end
        list.each_with_index do |prop, idx|
          b.open_tag("property", { fmtid: "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", pid: (idx + 2).to_s, name: prop[:name] })
          case prop[:type]
          when :number, :int, :i4
            b.tag("vt:i4") { |_| b.text(prop[:value].to_s) }
          when :float, :r8
            b.tag("vt:r8") { |_| b.text(prop[:value].to_s) }
          when :bool, :boolean
            b.tag("vt:bool") { |_| b.text(prop[:value] ? "true" : "false") }
          when :date, :filetime
            b.tag("vt:filetime") { |_| b.text(prop[:value].to_s) }
          else
            b.tag("vt:lpwstr") { |_| b.text(prop[:value].to_s) }
          end
          b.close_tag("property")
        end
        b.close_tag("Properties")
        io.string
      end

      def resolve_active_tab_index
        active_tab = @workbook_properties&.[](:active_sheet_index) || @workbook_properties&.[](:active_sheet) || @workbook_properties&.[](:active_tab)
        if active_tab.nil?
          first_vis_idx = @sheets.find_index do |s|
            state = s.respond_to?(:state) ? s.state : s[:state]
            state.nil? || state.to_sym == :visible
          end
          active_tab = first_vis_idx if first_vis_idx&.positive?
        end

        if active_tab.is_a?(String) || active_tab.is_a?(Symbol)
          target_name = active_tab.to_s
          idx = @sheets.find_index do |s|
            (s.respond_to?(:name) ? s.name : s[:name]) == target_name
          end
          raise ArgumentError, "Active sheet '#{target_name}' not found" if idx.nil?

          active_tab = idx
        end

        if active_tab
          target_sheet = @sheets[active_tab]
          raise ArgumentError, "Active sheet index #{active_tab} out of bounds" if target_sheet.nil?

          state = target_sheet.respond_to?(:state) ? target_sheet.state : target_sheet[:state]
          if state && state.to_sym != :visible
            name = target_sheet.respond_to?(:name) ? target_sheet.name : target_sheet[:name]
            raise ArgumentError, "Cannot set hidden sheet '#{name}' as active sheet"
          end
        end

        active_tab
      end

      def resolve_first_sheet_index
        first_sheet = @workbook_properties&.[](:first_sheet)
        return nil if first_sheet.nil?

        if first_sheet.is_a?(String) || first_sheet.is_a?(Symbol)
          target_name = first_sheet.to_s
          idx = @sheets.find_index do |s|
            (s.respond_to?(:name) ? s.name : s[:name]) == target_name
          end
          raise ArgumentError, "First sheet '#{target_name}' not found" if idx.nil?

          first_sheet = idx
        end

        if first_sheet
          target_sheet = @sheets[first_sheet]
          raise ArgumentError, "First sheet index #{first_sheet} out of bounds" if target_sheet.nil?
        end

        first_sheet
      end
    end
  end
end
