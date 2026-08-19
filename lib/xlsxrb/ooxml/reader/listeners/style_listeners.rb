# frozen_string_literal: true

# rbs_inline: enabled

require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  module Ooxml
    class Reader
      # SAX2 listener for parsing styles.xml (numFmts + cellXfs).
      class StylesListener
        include REXML::SAX2Listener

        attr_reader :num_fmts, :cell_xfs, :cell_style_xfs, :cell_styles, :fonts, :fills, :borders, :dxfs,
                    :indexed_colors, :mru_colors, :table_styles

        def initialize
          @num_fmts = {} # { numFmtId => formatCode }
          @cell_xfs = [] # Array of { num_fmt_id:, font_id:, fill_id:, border_id: }
          @cell_style_xfs = [] # Array of { num_fmt_id:, font_id:, fill_id:, border_id: }
          @cell_styles = [] # Array of { name:, xf_id:, builtin_id: }
          @fonts = []
          @fills = []
          @borders = []
          @dxfs = []
          @indexed_colors = []
          @mru_colors = []
          @table_styles = {}
          @inside_cell_xfs = false
          @inside_cell_style_xfs = false
          @inside_cell_styles = false
          @inside_fonts = false
          @inside_fills = false
          @inside_borders = false
          @inside_dxfs = false
          @inside_indexed_colors = false
          @inside_mru_colors = false
          @inside_table_styles = false
          @current_table_style = nil
          @current_font = nil
          @current_fill = nil
          @current_border = nil
          @current_border_side = nil
          @current_dxf = nil
          @current_xf = nil
          @current_gradient_stop = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)

          case name
          when "numFmt"
            id = attributes["numFmtId"]&.to_i
            code = attributes["formatCode"]
            if @inside_dxfs && @current_dxf
              @current_dxf[:num_fmt] = { num_fmt_id: id, format_code: code } if id && code
            elsif id && code
              @num_fmts[id] = code
            end
          when "cellXfs"
            @inside_cell_xfs = true
          when "cellStyleXfs"
            @inside_cell_style_xfs = true
          when "cellStyles"
            @inside_cell_styles = true
          when "cellStyle"
            if @inside_cell_styles
              entry = { name: attributes["name"] }
              entry[:xf_id] = attributes["xfId"].to_i if attributes["xfId"]
              entry[:builtin_id] = attributes["builtinId"].to_i if attributes["builtinId"]
              il = attributes["iLevel"]
              entry[:i_level] = il.to_i if il
              entry[:hidden] = true if %w[1 true].include?(attributes["hidden"])
              entry[:custom_builtin] = true if %w[1 true].include?(attributes["customBuiltin"])
              @cell_styles << entry
            end
          when "xf"
            xf_entry = {
              num_fmt_id: attributes["numFmtId"]&.to_i,
              font_id: attributes["fontId"]&.to_i,
              fill_id: attributes["fillId"]&.to_i,
              border_id: attributes["borderId"]&.to_i
            }
            %w[applyNumberFormat applyFont applyFill applyBorder applyAlignment applyProtection].each do |attr|
              key = attr.gsub(/[A-Z]/) { |m| "_#{m.downcase}" }.to_sym
              xf_entry[key] = attributes[attr] == "1" unless attributes[attr].nil?
            end
            if @inside_cell_xfs
              xf_entry[:xf_id] = attributes["xfId"]&.to_i
              xf_entry[:quote_prefix] = true if attributes["quotePrefix"] == "1"
              xf_entry[:pivot_button] = true if attributes["pivotButton"] == "1"
              @cell_xfs << xf_entry
              @current_xf = xf_entry
            elsif @inside_cell_style_xfs
              @cell_style_xfs << xf_entry
              @current_xf = xf_entry
            else
              @current_xf = nil
            end
          when "alignment"
            alignment = {}
            alignment[:horizontal] = attributes["horizontal"] if attributes["horizontal"]
            alignment[:vertical] = attributes["vertical"] if attributes["vertical"]
            alignment[:wrap_text] = true if attributes["wrapText"] == "1"
            alignment[:text_rotation] = attributes["textRotation"].to_i if attributes["textRotation"]
            alignment[:indent] = attributes["indent"].to_i if attributes["indent"]
            ri = attributes["relativeIndent"]
            alignment[:relative_indent] = ri.to_i if ri
            alignment[:shrink_to_fit] = true if attributes["shrinkToFit"] == "1"
            ro = attributes["readingOrder"]
            alignment[:reading_order] = ro.to_i if ro
            alignment[:justify_last_line] = true if attributes["justifyLastLine"] == "1"
            unless alignment.empty?
              if @inside_dxfs && @current_dxf
                @current_dxf[:alignment] = alignment
              elsif @current_xf
                @current_xf[:alignment] = alignment
              end
            end
          when "protection"
            protection = {}
            protection[:locked] = attributes["locked"] != "0" if attributes.key?("locked")
            protection[:hidden] = attributes["hidden"] == "1" if attributes.key?("hidden")
            unless protection.empty?
              if @inside_dxfs && @current_dxf
                @current_dxf[:protection] = protection
              elsif @current_xf
                @current_xf[:protection] = protection
              end
            end
          when "fonts"
            @inside_fonts = true
          when "font"
            @current_font = {} if @inside_fonts || @inside_dxfs
          when "b"
            @current_font[:bold] = true if @current_font
          when "i"
            @current_font[:italic] = true if @current_font
          when "strike"
            @current_font[:strike] = true if @current_font
          when "shadow"
            @current_font[:shadow] = true if @current_font
          when "outline"
            @current_font[:outline] = true if @current_font
          when "condense"
            @current_font[:condense] = true if @current_font
          when "extend"
            @current_font[:extend] = true if @current_font
          when "u"
            if @current_font
              val = attributes["val"]
              @current_font[:underline] = val || true
            end
          when "vertAlign"
            @current_font[:vert_align] = attributes["val"] if @current_font && attributes["val"]
          when "scheme"
            @current_font[:scheme] = attributes["val"] if @current_font && attributes["val"]
          when "family"
            @current_font[:family] = attributes["val"].to_i if @current_font && attributes["val"]
          when "charset"
            @current_font[:charset] = attributes["val"].to_i if @current_font && attributes["val"]
          when "sz"
            @current_font[:sz] = attributes["val"]&.to_f if @current_font
          when "color"
            parse_color(attributes)
          when "fills"
            @inside_fills = true
          when "fill"
            @current_fill = {} if @inside_fills || @inside_dxfs
          when "patternFill"
            @current_fill[:pattern] = attributes["patternType"] if @current_fill
          when "gradientFill"
            if @current_fill
              gradient = {}
              gradient[:type] = attributes["type"] if attributes["type"]
              gradient[:degree] = attributes["degree"].to_f if attributes["degree"]
              gradient[:left] = attributes["left"].to_f if attributes["left"]
              gradient[:right] = attributes["right"].to_f if attributes["right"]
              gradient[:top] = attributes["top"].to_f if attributes["top"]
              gradient[:bottom] = attributes["bottom"].to_f if attributes["bottom"]
              gradient[:stops] = []
              @current_fill[:gradient] = gradient
            end
          when "stop"
            @current_gradient_stop = { position: attributes["position"].to_f } if @current_fill&.dig(:gradient)
          when "fgColor"
            parse_fill_color(:fg_color, attributes) if @current_fill
          when "bgColor"
            parse_fill_color(:bg_color, attributes) if @current_fill
          when "borders"
            @inside_borders = true
          when "border"
            if @inside_borders || @inside_dxfs
              @current_border = {}
              @current_border[:diagonal_up] = true if attributes["diagonalUp"] == "1"
              @current_border[:diagonal_down] = true if attributes["diagonalDown"] == "1"
              ol = attributes["outline"]
              @current_border[:outline] = %w[1 true].include?(ol) unless ol.nil?
            end
          when "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"
            if @current_border
              style = attributes["style"]
              @current_border_side = name.to_sym
              @current_border[@current_border_side] = { style: style } if style
            end
          when "dxfs"
            @inside_dxfs = true
          when "dxf"
            @current_dxf = {}
          when "indexedColors"
            @inside_indexed_colors = true
          when "mruColors"
            @inside_mru_colors = true
          when "rgbColor"
            @indexed_colors << attributes["rgb"] if @inside_indexed_colors && attributes["rgb"]
          when "tableStyles"
            @inside_table_styles = true
            @table_styles[:default_table_style] = attributes["defaultTableStyle"] if attributes["defaultTableStyle"]
            @table_styles[:default_pivot_style] = attributes["defaultPivotStyle"] if attributes["defaultPivotStyle"]
            @table_styles[:styles] = []
          when "tableStyle"
            if @inside_table_styles
              ts = { name: attributes["name"], elements: [] }
              ts[:pivot] = %w[1 true].include?(attributes["pivot"]) if attributes.key?("pivot")
              ts[:table] = %w[1 true].include?(attributes["table"]) if attributes.key?("table")
              @current_table_style = ts
            end
          when "tableStyleElement"
            if @current_table_style
              el = { type: attributes["type"] }
              el[:size] = attributes["size"].to_i if attributes["size"]
              el[:dxf_id] = attributes["dxfId"].to_i if attributes["dxfId"]
              @current_table_style[:elements] << el
            end
          end

          parse_font_name(name, attributes)
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "cellXfs"
            @inside_cell_xfs = false
          when "cellStyleXfs"
            @inside_cell_style_xfs = false
          when "xf"
            @current_xf = nil
          when "cellStyles"
            @inside_cell_styles = false
          when "fonts"
            @inside_fonts = false
          when "font"
            if @inside_dxfs && @current_dxf
              @current_dxf[:font] = @current_font
            elsif @inside_fonts
              @fonts << @current_font
            end
            @current_font = nil
          when "fills"
            @inside_fills = false
          when "fill"
            if @inside_dxfs && @current_dxf
              @current_dxf[:fill] = @current_fill
            elsif @inside_fills
              @fills << @current_fill
            end
            @current_fill = nil
          when "stop"
            @current_fill[:gradient][:stops] << @current_gradient_stop if @current_gradient_stop && @current_fill&.dig(:gradient)
            @current_gradient_stop = nil
          when "borders"
            @inside_borders = false
          when "border"
            if @inside_dxfs && @current_dxf
              @current_dxf[:border] = @current_border
            elsif @inside_borders
              @borders << @current_border
            end
            @current_border = nil
          when "left", "right", "top", "bottom", "diagonal", "vertical", "horizontal", "start", "end"
            @current_border_side = nil
          when "dxfs"
            @inside_dxfs = false
          when "dxf"
            @dxfs << @current_dxf if @current_dxf
            @current_dxf = nil
          when "indexedColors"
            @inside_indexed_colors = false
          when "mruColors"
            @inside_mru_colors = false
          when "tableStyles"
            @inside_table_styles = false
          when "tableStyle"
            if @inside_table_styles && @current_table_style
              @table_styles[:styles] << @current_table_style
              @current_table_style = nil
            end
          end
        end

        private

        def parse_color(attributes)
          if @inside_mru_colors
            c = {}
            c[:auto] = true if %w[1 true].include?(attributes["auto"])
            c[:rgb] = attributes["rgb"] if attributes["rgb"]
            c[:theme] = attributes["theme"].to_i if attributes["theme"]
            c[:tint] = attributes["tint"].to_f if attributes["tint"]
            c[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
            @mru_colors << c unless c.empty?
          elsif @current_gradient_stop
            @current_gradient_stop[:auto] = true if %w[1 true].include?(attributes["auto"])
            @current_gradient_stop[:color] = attributes["rgb"] if attributes["rgb"]
            @current_gradient_stop[:theme] = attributes["theme"].to_i if attributes["theme"]
            @current_gradient_stop[:tint] = attributes["tint"].to_f if attributes["tint"]
            @current_gradient_stop[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
          elsif @current_border_side && @current_border
            side_data = @current_border[@current_border_side]
            if side_data.is_a?(Hash)
              side_data[:auto] = true if %w[1 true].include?(attributes["auto"])
              side_data[:color] = attributes["rgb"] if attributes["rgb"]
              side_data[:theme] = attributes["theme"].to_i if attributes["theme"]
              side_data[:tint] = attributes["tint"].to_f if attributes["tint"]
              side_data[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
            end
          elsif @current_font
            @current_font[:auto] = true if %w[1 true].include?(attributes["auto"])
            @current_font[:color] = attributes["rgb"] if attributes["rgb"]
            @current_font[:theme] = attributes["theme"].to_i if attributes["theme"]
            @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
            @current_font[:indexed] = attributes["indexed"].to_i if attributes["indexed"]
          end
        end

        def parse_fill_color(key, attributes)
          if %w[1 true].include?(attributes["auto"])
            @current_fill[:"#{key}_auto"] = true
          elsif attributes["rgb"]
            @current_fill[key] = attributes["rgb"]
          elsif attributes["theme"]
            @current_fill[:"#{key}_theme"] = attributes["theme"].to_i
            @current_fill[:"#{key}_tint"] = attributes["tint"].to_f if attributes["tint"]
          elsif attributes["indexed"]
            @current_fill[:"#{key}_indexed"] = attributes["indexed"].to_i
          end
        end

        def parse_font_name(name, attributes)
          @current_font[:name] = attributes["val"] if name == "name" && @current_font && attributes["val"]
        end

        def element_name(local_name, qname)
          if local_name.nil? || local_name.empty?
            qname.to_s.split(":").last
          else
            local_name
          end
        end
      end

      # SAX2 listener that captures cell style index (s attribute) from worksheet.
      class CellStyleListener
        include REXML::SAX2Listener

        attr_reader :cell_style_indices

        def initialize
          @cell_style_indices = {}
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "c"

          ref = attributes["r"]
          s = attributes["s"]
          @cell_style_indices[ref] = s.to_i if ref && s
        end

        private

        def element_name(local_name, qname)
          if local_name.nil? || local_name.empty?
            qname.to_s.split(":").last
          else
            local_name
          end
        end
      end

      # SAX2 listener for parsing <phoneticPr> element.
      class PhoneticPrListener
        include REXML::SAX2Listener

        attr_reader :phonetic_pr

        def initialize
          @phonetic_pr = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          return unless name == "phoneticPr"

          pr = {}
          pr[:font_id] = attributes["fontId"].to_i if attributes["fontId"]
          pr[:type] = attributes["type"] if attributes["type"]
          pr[:alignment] = attributes["alignment"] if attributes["alignment"]
          @phonetic_pr = pr
        end

        private

        def element_name(local_name, qname)
          if local_name.nil? || local_name.empty?
            qname.to_s.split(":").last
          else
            local_name
          end
        end
      end
    end
  end
end
