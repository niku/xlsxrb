# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    class Writer
      # Mixin containing Stylesheet XML generation logic.
      module StylesXml
        # : () -> untyped
        def generate_styles_xml
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("styleSheet", { xmlns: SSML_NS }) do
              # numFmts
              unless @num_fmts.empty?
                xml.tag("numFmts", { count: @num_fmts.size }) do
                  @num_fmts.each do |nf|
                    xml.empty_tag("numFmt", { numFmtId: nf[:num_fmt_id] || "", formatCode: nf[:format_code] || "" })
                  end
                end
              end

              # fonts
              xml.tag("fonts", { count: @fonts.size }) do
                @fonts.each { |f| write_font_xml(xml, f) }
              end

              # fills
              xml.tag("fills", { count: @fills.size }) do
                @fills.each { |f| write_fill_xml(xml, f) }
              end

              # borders
              xml.tag("borders", { count: @borders.size }) do
                @borders.each { |bdr| write_border_xml(xml, bdr) }
              end

              # cellStyleXfs
              xml.tag("cellStyleXfs", { count: @cell_style_xfs.size }) do
                @cell_style_xfs.each do |xf|
                  xml.empty_tag("xf", {
                                  numFmtId: xf[:num_fmt_id] || "",
                                  fontId: xf[:font_id] || "",
                                  fillId: xf[:fill_id] || "",
                                  borderId: xf[:border_id] || ""
                                })
                end
              end

              # cellXfs
              xml.tag("cellXfs", { count: @xf_entries.size }) do
                @xf_entries.each do |xf|
                  xf_attrs = {
                    numFmtId: xf[:num_fmt_id] || "",
                    fontId: xf[:font_id] || "",
                    fillId: xf[:fill_id] || "",
                    borderId: xf[:border_id] || "",
                    xfId: xf[:xf_id] || 0,
                    applyNumberFormat: xf[:num_fmt_id].positive? ? 1 : nil,
                    applyFont: xf[:font_id].positive? ? 1 : nil,
                    applyFill: xf[:fill_id].positive? ? 1 : nil,
                    applyBorder: xf[:border_id].positive? ? 1 : nil,
                    applyAlignment: xf[:alignment] ? 1 : nil,
                    applyProtection: xf[:protection] ? 1 : nil,
                    quotePrefix: xf[:quote_prefix] ? 1 : nil,
                    pivotButton: xf[:pivot_button] ? 1 : nil
                  }
                  if xf[:alignment] || xf[:protection]
                    xml.tag("xf", xf_attrs) do
                      write_alignment_xml(xml, xf[:alignment]) if xf[:alignment]
                      write_protection_xml(xml, xf[:protection]) if xf[:protection]
                    end
                  else
                    xml.empty_tag("xf", xf_attrs)
                  end
                end
              end

              # cellStyles
              xml.tag("cellStyles", { count: @cell_style_names.size }) do
                @cell_style_names.each do |cs|
                  xml.empty_tag("cellStyle", {
                                  name: cs[:name] || "",
                                  xfId: cs[:xf_id] || "",
                                  builtinId: cs[:builtin_id],
                                  iLevel: cs[:i_level],
                                  hidden: cs[:hidden] ? 1 : nil,
                                  customBuiltin: cs[:custom_builtin] ? 1 : nil
                                })
                end
              end

              # dxfs
              unless @dxfs.empty?
                xml.tag("dxfs", { count: @dxfs.size }) do
                  @dxfs.each { |d| write_dxf_xml(xml, d) }
                end
              end

              # tableStyles
              ts_styles = @table_styles[:styles] || []
              unless ts_styles.empty? && @table_styles[:default_table_style].nil? && @table_styles[:default_pivot_style].nil?
                ts_attrs = {
                  count: ts_styles.size,
                  defaultTableStyle: @table_styles[:default_table_style],
                  defaultPivotStyle: @table_styles[:default_pivot_style]
                }
                if ts_styles.empty?
                  xml.empty_tag("tableStyles", ts_attrs)
                else
                  xml.tag("tableStyles", ts_attrs) do
                    ts_styles.each do |ts|
                      s_attrs = {
                        name: ts[:name] || "",
                        pivot: ts[:pivot] == false ? 0 : nil,
                        table: ts[:table] == false ? 0 : nil,
                        count: ts[:elements].empty? ? nil : ts[:elements].size
                      }
                      if ts[:elements].empty?
                        xml.empty_tag("tableStyle", s_attrs)
                      else
                        xml.tag("tableStyle", s_attrs) do
                          ts[:elements].each do |el|
                            xml.empty_tag("tableStyleElement", {
                                            type: el[:type] || "",
                                            size: el[:size] && el[:size] != 1 ? el[:size] : nil,
                                            dxfId: el[:dxf_id]
                                          })
                          end
                        end
                      end
                    end
                  end
                end
              end

              # colors
              unless @indexed_colors.empty? && @mru_colors.empty?
                xml.tag("colors") do
                  unless @indexed_colors.empty?
                    xml.tag("indexedColors") do
                      @indexed_colors.each do |c|
                        xml.empty_tag("rgbColor", { rgb: format_rgb_hex(c) || c || "" })
                      end
                    end
                  end
                  unless @mru_colors.empty?
                    xml.tag("mruColors") do
                      @mru_colors.each { |c| write_color_xml(xml, c) }
                    end
                  end
                end
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped alignment) -> void
        def write_alignment_xml(xml, alignment)
          xml.empty_tag("alignment", {
                          horizontal: alignment[:horizontal],
                          vertical: alignment[:vertical],
                          wrapText: alignment[:wrap_text] ? 1 : nil,
                          textRotation: alignment[:text_rotation],
                          indent: alignment[:indent],
                          relativeIndent: alignment[:relative_indent],
                          shrinkToFit: alignment[:shrink_to_fit] ? 1 : nil,
                          readingOrder: alignment[:reading_order],
                          justifyLastLine: alignment[:justify_last_line] ? 1 : nil
                        })
        end

        # : (XmlBuilder xml, untyped protection) -> void
        def write_protection_xml(xml, protection)
          xml.empty_tag("protection", {
                          locked: xml_bool_attr(protection[:locked]),
                          hidden: xml_bool_attr(protection[:hidden])
                        })
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

        # : (XmlBuilder xml, untyped source, ?tag: ::String) -> void
        def write_color_xml(xml, source, tag: "color")
          return unless source

          spec = if source.is_a?(Hash) && source[:color]
                   source[:color]
                 else
                   source
                 end
          return unless spec

          if spec.is_a?(Hash)
            if spec[:auto]
              xml.empty_tag(tag, { auto: 1 })
            elsif spec[:rgb] || (spec[:color] && !spec[:color].is_a?(Hash))
              formatted_rgb = format_rgb_hex(spec[:rgb] || spec[:color])
              xml.empty_tag(tag, { rgb: formatted_rgb }) if formatted_rgb
            elsif spec[:theme]
              xml.empty_tag(tag, { theme: spec[:theme], tint: spec[:tint] })
            elsif spec[:indexed]
              xml.empty_tag(tag, { indexed: spec[:indexed] })
            end
          elsif spec.is_a?(String) || spec.is_a?(Symbol)
            formatted_rgb = format_rgb_hex(spec)
            xml.empty_tag(tag, { rgb: formatted_rgb }) if formatted_rgb
          end
        end

        # : (XmlBuilder xml, untyped font) -> void
        def write_font_xml(xml, font)
          xml.tag("font") do
            xml.empty_tag("b", font[:bold] ? XmlBuilder::EMPTY_HASH : { val: 0 }) if font.key?(:bold)
            xml.empty_tag("i", font[:italic] ? XmlBuilder::EMPTY_HASH : { val: 0 }) if font.key?(:italic)
            xml.empty_tag("strike", font[:strike] ? XmlBuilder::EMPTY_HASH : { val: 0 }) if font.key?(:strike)
            xml.empty_tag("shadow") if font[:shadow]
            xml.empty_tag("outline") if font[:outline]
            xml.empty_tag("condense") if font[:condense]
            xml.empty_tag("extend") if font[:extend]
            if font[:underline]
              u_val = font[:underline]
              if u_val == true || u_val.to_s == "single"
                xml.empty_tag("u")
              else
                mapped_u = StyleBuilder::UNDERLINE_MAP[u_val.to_sym] || u_val
                xml.empty_tag("u", { val: mapped_u })
              end
            end
            va_val = font[:vert_align] || font[:vertical_align]
            va_val ||= "superscript" if font[:superscript]
            va_val ||= "subscript" if font[:subscript]
            if va_val
              mapped_va = StyleBuilder::VERT_ALIGN_MAP[va_val.to_sym] || va_val
              xml.empty_tag("vertAlign", { val: mapped_va })
            end
            font_sz = font[:size] || font[:sz]
            xml.empty_tag("sz", { val: font_sz }) if font_sz
            write_color_xml(xml, font)
            xml.empty_tag("name", { val: font[:name] }) if font[:name]
            xml.empty_tag("family", { val: font[:family] }) if font[:family]
            xml.empty_tag("charset", { val: font[:charset] }) if font[:charset]
            xml.empty_tag("scheme", { val: font[:scheme] }) if font[:scheme]
          end
        end

        # : (XmlBuilder xml, untyped fill) -> void
        def write_fill_xml(xml, fill)
          return write_gradient_fill_xml(xml, fill[:gradient]) if fill[:gradient]

          has_fg = fill[:fg_color] || fill[:fg_color_theme] || fill[:fg_color_indexed] || fill[:fg_color_auto]
          has_bg = fill[:bg_color] || fill[:bg_color_theme] || fill[:bg_color_indexed] || fill[:bg_color_auto]
          fill_pattern = fill[:pattern_type] || fill[:pattern]
          if fill_pattern && !has_fg && !has_bg
            xml.tag("fill") { xml.empty_tag("patternFill", { patternType: fill_pattern }) }
            return
          end

          xml.tag("fill") do
            xml.tag("patternFill", { patternType: fill_pattern || "solid" }) do
              write_fill_color_xml(xml, "fgColor", fill, :fg)
              write_fill_color_xml(xml, "bgColor", fill, :bg)
            end
          end
        end

        # : (XmlBuilder xml, untyped tag, untyped fill, untyped prefix) -> void
        def write_fill_color_xml(xml, tag, fill, prefix)
          val = fill[:"#{prefix}_color"]
          if val
            write_color_xml(xml, val, tag: tag)
          elsif fill[:"#{prefix}_color_theme"]
            xml.empty_tag(tag, { theme: fill[:"#{prefix}_color_theme"], tint: fill[:"#{prefix}_color_tint"] })
          elsif fill[:"#{prefix}_color_indexed"]
            xml.empty_tag(tag, { indexed: fill[:"#{prefix}_color_indexed"] })
          elsif fill[:"#{prefix}_color_auto"]
            xml.empty_tag(tag, { auto: 1 })
          end
        end

        # : (XmlBuilder xml, untyped gradient) -> void
        def write_gradient_fill_xml(xml, gradient)
          attrs = {
            type: gradient[:type],
            degree: gradient[:degree],
            left: gradient[:left],
            right: gradient[:right],
            top: gradient[:top],
            bottom: gradient[:bottom]
          }
          xml.tag("fill") do
            if gradient[:stops]&.any?
              xml.tag("gradientFill", attrs) do
                gradient[:stops].each do |stop|
                  xml.tag("stop", { position: stop[:position] || "" }) { write_color_xml(xml, stop) }
                end
              end
            else
              xml.empty_tag("gradientFill", attrs)
            end
          end
        end

        # : (XmlBuilder xml, untyped brk, default_max: untyped) -> void
        def write_brk_xml(xml, brk, default_max:)
          brk = brk.to_h if brk.is_a?(Elements::MetadataRecord)
          if brk.is_a?(Hash)
            max_val = brk.respond_to?(:fetch) ? brk.fetch(:max, default_max) : (brk[:max] || default_max)
            man_val = if brk.respond_to?(:fetch)
                        brk.fetch(:man, true)
                      else
                        brk[:man].nil? || brk[:man]
                      end
            xml.empty_tag("brk", {
                            id: brk[:id] || "",
                            min: brk[:min],
                            max: max_val || "",
                            man: man_val ? 1 : nil,
                            pt: brk[:pt] ? 1 : nil
                          })
          else
            xml.empty_tag("brk", { id: brk || "", max: default_max || "", man: 1 })
          end
        end

        # : (XmlBuilder xml, untyped bdr) -> void
        def write_border_xml(xml, bdr)
          b_attrs = {
            diagonalUp: bdr[:diagonal_up] ? 1 : nil,
            diagonalDown: bdr[:diagonal_down] ? 1 : nil,
            outline: bdr[:outline] == false ? 0 : nil
          }
          xml.tag("border", b_attrs) do
            %i[left right top bottom diagonal vertical horizontal].each do |side|
              s = bdr[side]
              if s.is_a?(Hash)
                xml.tag(side.to_s, { style: s[:style] || "" }) { write_color_xml(xml, s) }
              else
                xml.empty_tag(side.to_s)
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped dxf) -> void
        def write_dxf_xml(xml, dxf)
          dxf = dxf.to_h if dxf.is_a?(Elements::MetadataRecord)
          xml.tag("dxf") do
            write_font_xml(xml, dxf[:font]) if dxf[:font]
            nf = dxf[:number_format] || dxf[:num_fmt]
            if nf
              nf_h = nf.is_a?(Elements::MetadataRecord) ? nf.to_h : nf
              xml.empty_tag("numFmt", { numFmtId: nf_h[:num_fmt_id] || "", formatCode: nf_h[:format_code] || "" })
            end
            write_fill_xml(xml, dxf[:fill]) if dxf[:fill]
            write_alignment_xml(xml, dxf[:alignment]) if dxf[:alignment]
            write_border_xml(xml, dxf[:border]) if dxf[:border]
            write_protection_xml(xml, dxf[:protection]) if dxf[:protection]
          end
        end

        # : (untyped cell_address) -> untyped
        def validate_cell_address!(cell_address)
          raise ArgumentError, "cell address must be a String" unless cell_address.is_a?(String)

          match = cell_address.match(CELL_ADDRESS_PATTERN)
          raise ArgumentError, "invalid cell address: #{cell_address.inspect}" unless match

          row_num = match[2].to_i
          raise ArgumentError, "row out of range: #{row_num}" unless row_num.between?(1, MAX_ROW)

          col_index = column_letter_to_index(match[1])
          raise ArgumentError, "column out of range: #{match[1]}" unless col_index.between?(1, MAX_COLUMN_INDEX)
        end

        # : (untyped letters) -> untyped
        def column_letter_to_index(letters)
          letters.chars.reduce(0) { |sum, char| (sum * 26) + (char.ord - "A".ord + 1) }
        end

        # Extracts the column letter(s) from a cell address, e.g. "A" from "A1".
        # : (untyped cell_address) -> untyped
        def extract_column_letter(cell_address)
          cell_address.match(/^([A-Z]+)/)[1]
        end
      end
    end
  end
end
