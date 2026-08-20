# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    class Writer
      # Mixin containing Stylesheet XML generation logic.
      module StylesXml
        # : () -> untyped
        def generate_styles_xml
          parts = [
            XML_HEADER,
            %(<styleSheet xmlns="#{SSML_NS}">)
          ]

          # numFmts
          unless @num_fmts.empty?
            parts << %(<numFmts count="#{@num_fmts.size}">)
            @num_fmts.each do |nf|
              parts << %(<numFmt numFmtId="#{nf[:num_fmt_id]}" formatCode="#{xml_escape(nf[:format_code])}"/>)
            end
            parts << "</numFmts>"
          end

          # fonts
          parts << %(<fonts count="#{@fonts.size}">)
          @fonts.each { |f| parts << emit_font_xml(f) }
          parts << "</fonts>"

          # fills
          parts << %(<fills count="#{@fills.size}">)
          @fills.each { |f| parts << emit_fill_xml(f) }
          parts << "</fills>"

          # borders
          parts << %(<borders count="#{@borders.size}">)
          @borders.each { |b| parts << emit_border_xml(b) }
          parts << "</borders>"

          # cellStyleXfs
          parts << %(<cellStyleXfs count="#{@cell_style_xfs.size}">)
          @cell_style_xfs.each do |xf|
            parts << %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}"/>)
          end
          parts << "</cellStyleXfs>"

          # cellXfs
          parts << %(<cellXfs count="#{@xf_entries.size}">)
          @xf_entries.each do |xf|
            apply_attrs = []
            apply_attrs << ' applyNumberFormat="1"' if xf[:num_fmt_id].positive?
            apply_attrs << ' applyFont="1"' if xf[:font_id].positive?
            apply_attrs << ' applyFill="1"' if xf[:fill_id].positive?
            apply_attrs << ' applyBorder="1"' if xf[:border_id].positive?
            apply_attrs << ' applyAlignment="1"' if xf[:alignment]
            apply_attrs << ' applyProtection="1"' if xf[:protection]
            apply_attrs << ' quotePrefix="1"' if xf[:quote_prefix]
            apply_attrs << ' pivotButton="1"' if xf[:pivot_button]
            children = []
            children << emit_alignment_xml(xf[:alignment]) if xf[:alignment]
            children << emit_protection_xml(xf[:protection]) if xf[:protection]
            xf_id = xf[:xf_id] || 0
            parts << if children.any?
                       %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}" xfId="#{xf_id}"#{apply_attrs.join}>#{children.join}</xf>)
                     else
                       %(<xf numFmtId="#{xf[:num_fmt_id]}" fontId="#{xf[:font_id]}" fillId="#{xf[:fill_id]}" borderId="#{xf[:border_id]}" xfId="#{xf_id}"#{apply_attrs.join}/>)
                     end
          end
          parts << "</cellXfs>"

          # cellStyles
          parts << %(<cellStyles count="#{@cell_style_names.size}">)
          @cell_style_names.each do |cs|
            cs_attrs = %(name="#{xml_escape(cs[:name])}" xfId="#{cs[:xf_id]}")
            cs_attrs << %( builtinId="#{cs[:builtin_id]}") if cs[:builtin_id]
            cs_attrs << %( iLevel="#{cs[:i_level]}") if cs[:i_level]
            cs_attrs << ' hidden="1"' if cs[:hidden]
            cs_attrs << ' customBuiltin="1"' if cs[:custom_builtin]
            parts << "<cellStyle #{cs_attrs}/>"
          end
          parts << "</cellStyles>"

          # dxfs
          unless @dxfs.empty?
            parts << %(<dxfs count="#{@dxfs.size}">)
            @dxfs.each { |d| parts << emit_dxf_xml(d) }
            parts << "</dxfs>"
          end

          # tableStyles
          ts_styles = @table_styles[:styles] || []
          unless ts_styles.empty? && @table_styles[:default_table_style].nil? && @table_styles[:default_pivot_style].nil?
            ts_attrs = [%(count="#{ts_styles.size}")]
            ts_attrs << %(defaultTableStyle="#{xml_escape(@table_styles[:default_table_style])}") if @table_styles[:default_table_style]
            ts_attrs << %(defaultPivotStyle="#{xml_escape(@table_styles[:default_pivot_style])}") if @table_styles[:default_pivot_style]
            if ts_styles.empty?
              parts << "<tableStyles #{ts_attrs.join(" ")}/>"
            else
              parts << "<tableStyles #{ts_attrs.join(" ")}>"
              ts_styles.each do |ts|
                s_attrs = [%(name="#{xml_escape(ts[:name])}")]
                s_attrs << %(pivot="0") if ts[:pivot] == false
                s_attrs << %(table="0") if ts[:table] == false
                s_attrs << %(count="#{ts[:elements].size}") unless ts[:elements].empty?
                if ts[:elements].empty?
                  parts << "<tableStyle #{s_attrs.join(" ")}/>"
                else
                  parts << "<tableStyle #{s_attrs.join(" ")}>"
                  ts[:elements].each do |el|
                    el_attrs = [%(type="#{el[:type]}")]
                    el_attrs << %(size="#{el[:size]}") if el[:size] && el[:size] != 1
                    el_attrs << %(dxfId="#{el[:dxf_id]}") if el[:dxf_id]
                    parts << "<tableStyleElement #{el_attrs.join(" ")}/>"
                  end
                  parts << "</tableStyle>"
                end
              end
              parts << "</tableStyles>"
            end
          end

          # colors
          unless @indexed_colors.empty? && @mru_colors.empty?
            parts << "<colors>"
            unless @indexed_colors.empty?
              parts << "<indexedColors>"
              @indexed_colors.each { |c| parts << %(<rgbColor rgb="#{c}"/>) }
              parts << "</indexedColors>"
            end
            unless @mru_colors.empty?
              parts << "<mruColors>"
              @mru_colors.each { |c| parts << emit_color_xml(c) }
              parts << "</mruColors>"
            end
            parts << "</colors>"
          end

          parts << "</styleSheet>"
          parts.join
        end

        # : (untyped alignment) -> ::String
        def emit_alignment_xml(alignment)
          attrs = []
          attrs << %(horizontal="#{alignment[:horizontal]}") if alignment[:horizontal]
          attrs << %(vertical="#{alignment[:vertical]}") if alignment[:vertical]
          attrs << %(wrapText="1") if alignment[:wrap_text]
          attrs << %(textRotation="#{alignment[:text_rotation]}") if alignment[:text_rotation]
          attrs << %(indent="#{alignment[:indent]}") if alignment[:indent]
          attrs << %(relativeIndent="#{alignment[:relative_indent]}") if alignment[:relative_indent]
          attrs << %(shrinkToFit="1") if alignment[:shrink_to_fit]
          attrs << %(readingOrder="#{alignment[:reading_order]}") if alignment[:reading_order]
          attrs << %(justifyLastLine="1") if alignment[:justify_last_line]
          "<alignment #{attrs.join(" ")}/>"
        end

        # : (untyped protection) -> ::String
        def emit_protection_xml(protection)
          attrs = []
          attrs << %(locked="#{protection[:locked] ? "1" : "0"}") unless protection[:locked].nil?
          attrs << %(hidden="#{protection[:hidden] ? "1" : "0"}") unless protection[:hidden].nil?
          "<protection #{attrs.join(" ")}/>"
        end

        # : (untyped source, ?tag: ::String) -> (::String | ::String | ::String | ::String | "")
        def emit_color_xml(source, tag: "color")
          if source[:auto]
            %(<#{tag} auto="1"/>)
          elsif source[:color] || source[:rgb]
            %(<#{tag} rgb="#{source[:color] || source[:rgb]}"/>)
          elsif source[:theme]
            attrs = [%(theme="#{source[:theme]}")]
            attrs << %(tint="#{source[:tint]}") if source[:tint]
            %(<#{tag} #{attrs.join(" ")}/>)
          elsif source[:indexed]
            %(<#{tag} indexed="#{source[:indexed]}"/>)
          else
            ""
          end
        end

        # : (untyped font) -> untyped
        def emit_font_xml(font)
          parts = ["<font>"]
          parts << "<b/>" if font[:bold]
          parts << "<i/>" if font[:italic]
          parts << "<strike/>" if font[:strike]
          parts << "<shadow/>" if font[:shadow]
          parts << "<outline/>" if font[:outline]
          parts << "<condense/>" if font[:condense]
          parts << "<extend/>" if font[:extend]
          if font[:underline]
            parts << if font[:underline] == true
                       "<u/>"
                     else
                       %(<u val="#{font[:underline]}"/>)
                     end
          end
          parts << %(<vertAlign val="#{font[:vert_align]}"/>) if font[:vert_align]
          parts << %(<sz val="#{font[:sz]}"/>) if font[:sz]
          parts << emit_color_xml(font)
          parts << %(<name val="#{xml_escape(font[:name])}"/>) if font[:name]
          parts << %(<family val="#{font[:family]}"/>) if font[:family]
          parts << %(<charset val="#{font[:charset]}"/>) if font[:charset]
          parts << %(<scheme val="#{font[:scheme]}"/>) if font[:scheme]
          parts << "</font>"
          parts.join
        end

        # : (untyped fill) -> (untyped | ::String)
        def emit_fill_xml(fill)
          return emit_gradient_fill_xml(fill[:gradient]) if fill[:gradient]

          has_fg = fill[:fg_color] || fill[:fg_color_theme] || fill[:fg_color_indexed] || fill[:fg_color_auto]
          has_bg = fill[:bg_color] || fill[:bg_color_theme] || fill[:bg_color_indexed] || fill[:bg_color_auto]
          return "<fill><patternFill patternType=\"#{fill[:pattern]}\"/></fill>" if fill[:pattern] && !has_fg && !has_bg

          parts = ["<fill>"]
          pt = fill[:pattern] || "solid"
          parts << %(<patternFill patternType="#{pt}">)
          parts << emit_fill_color_xml("fgColor", fill, :fg)
          parts << emit_fill_color_xml("bgColor", fill, :bg)
          parts << "</patternFill>"
          parts << "</fill>"
          parts.join
        end

        # : (untyped tag, untyped fill, untyped prefix) -> (::String | ::String | ::String | ::String | "")
        def emit_fill_color_xml(tag, fill, prefix)
          if fill[:"#{prefix}_color"]
            %(<#{tag} rgb="#{fill[:"#{prefix}_color"]}"/>)
          elsif fill[:"#{prefix}_color_theme"]
            attrs = [%(theme="#{fill[:"#{prefix}_color_theme"]}")]
            attrs << %(tint="#{fill[:"#{prefix}_color_tint"]}") if fill[:"#{prefix}_color_tint"]
            %(<#{tag} #{attrs.join(" ")}/>)
          elsif fill[:"#{prefix}_color_indexed"]
            %(<#{tag} indexed="#{fill[:"#{prefix}_color_indexed"]}"/>)
          elsif fill[:"#{prefix}_color_auto"]
            %(<#{tag} auto="1"/>)
          else
            ""
          end
        end

        # : (untyped gradient) -> untyped
        def emit_gradient_fill_xml(gradient)
          attrs = []
          attrs << %(type="#{gradient[:type]}") if gradient[:type]
          attrs << %(degree="#{gradient[:degree]}") if gradient[:degree]
          attrs << %(left="#{gradient[:left]}") if gradient[:left]
          attrs << %(right="#{gradient[:right]}") if gradient[:right]
          attrs << %(top="#{gradient[:top]}") if gradient[:top]
          attrs << %(bottom="#{gradient[:bottom]}") if gradient[:bottom]
          parts = ["<fill>"]
          parts << "<gradientFill#{" #{attrs.join(" ")}" unless attrs.empty?}"
          if gradient[:stops]&.any?
            parts[-1] = "#{parts[-1]}>"
            gradient[:stops].each do |stop|
              parts << %(<stop position="#{stop[:position]}">#{emit_color_xml(stop)}</stop>)
            end
            parts << "</gradientFill>"
          else
            parts[-1] = "#{parts[-1]}/>"
          end
          parts << "</fill>"
          parts.join
        end

        # : (untyped brk, default_max: untyped) -> ::String
        def emit_brk_xml(brk, default_max:)
          if brk.is_a?(Hash)
            attrs = %(id="#{brk[:id]}")
            attrs << %( min="#{brk[:min]}") if brk[:min]
            attrs << %( max="#{brk.fetch(:max, default_max)}")
            attrs << ' man="1"' if brk.fetch(:man, true)
            attrs << ' pt="1"' if brk[:pt]
            "<brk #{attrs}/>"
          else
            %(<brk id="#{brk}" max="#{default_max}" man="1"/>)
          end
        end

        # : (untyped bdr) -> untyped
        def emit_border_xml(bdr)
          border_attrs = []
          border_attrs << ' diagonalUp="1"' if bdr[:diagonal_up]
          border_attrs << ' diagonalDown="1"' if bdr[:diagonal_down]
          border_attrs << ' outline="0"' if bdr[:outline] == false
          parts = ["<border#{border_attrs.join}>"]
          %i[left right top bottom diagonal vertical horizontal].each do |side|
            s = bdr[side]
            if s.is_a?(Hash)
              parts << %(<#{side} style="#{s[:style]}">)
              parts << emit_color_xml(s)
              parts << "</#{side}>"
            else
              parts << "<#{side}/>"
            end
          end
          parts << "</border>"
          parts.join
        end

        # : (untyped dxf) -> untyped
        def emit_dxf_xml(dxf)
          parts = ["<dxf>"]
          parts << emit_font_xml(dxf[:font]) if dxf[:font]
          if dxf[:num_fmt]
            nf = dxf[:num_fmt]
            parts << %(<numFmt numFmtId="#{nf[:num_fmt_id]}" formatCode="#{xml_escape(nf[:format_code])}"/>)
          end
          parts << emit_fill_xml(dxf[:fill]) if dxf[:fill]
          parts << emit_alignment_xml(dxf[:alignment]) if dxf[:alignment]
          parts << emit_border_xml(dxf[:border]) if dxf[:border]
          parts << emit_protection_xml(dxf[:protection]) if dxf[:protection]
          parts << "</dxf>"
          parts.join
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
