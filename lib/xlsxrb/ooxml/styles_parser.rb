# frozen_string_literal: true

# rbs_inline: enabled

require_relative "xml_parser"
require_relative "../elements/styles"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/styles.xml.
    # Returns an Elements::Styles (Hash) with :num_fmts, :fonts, :fills, :borders, :cell_xfs, :cell_style_xfs.
    class StylesParser
      #: (String? xml_string) -> Elements::Styles
      def self.parse(xml_string)
        return Elements::Styles.new if xml_string.nil? || xml_string.empty?

        listener = Listener.new
        XmlParser.parse(xml_string, listener)
        styles = Elements::Styles.new(listener.result)
        styles.precompute!
        styles
      end

      # SAX listener for parsing styles.xml content.
      class Listener
        include REXML::SAX2Listener

        attr_reader :result

        def initialize
          @result = {
            num_fmts: {},
            fonts: [],
            fills: [],
            borders: [],
            cell_xfs: [],
            cell_style_xfs: [],
            dxfs: []
          }
          @context = []
          @current_font = nil
          @current_fill = nil
          @current_fill_pattern = nil
          @current_border = nil
          @current_border_side = nil
          @current_xf = nil
          @in_cell_xfs = false
          @in_cell_style_xfs = false
          @current_dxf = nil
        end

        def start_element(_uri, localname, _qname, attrs)
          @context.push(localname)
          case localname
          when "numFmt"
            id = attrs["numFmtId"]&.to_i
            code = attrs["formatCode"]
            if @current_dxf
              @current_dxf[:num_fmt] = { num_fmt_id: id, format_code: code } if id && code
            elsif parent_context?("numFmts") && id && code
              @result[:num_fmts][id] = code
            end
          when "dxf"
            @current_dxf = {} if parent_context?("dxfs")
          when "font"
            @current_font = {} if parent_context?("fonts") || (@current_dxf && parent_context?("dxf"))
          when "b"
            @current_font[:bold] = truthy_val?(attrs) if @current_font
          when "i"
            @current_font[:italic] = truthy_val?(attrs) if @current_font
          when "strike"
            @current_font[:strike] = truthy_val?(attrs) if @current_font
          when "shadow"
            @current_font[:shadow] = truthy_val?(attrs) if @current_font
          when "outline"
            @current_font[:outline] = truthy_val?(attrs) if @current_font
          when "condense"
            @current_font[:condense] = truthy_val?(attrs) if @current_font
          when "extend"
            @current_font[:extend] = truthy_val?(attrs) if @current_font
          when "u"
            @current_font[:underline] = attrs["val"] || "single" if @current_font
          when "sz"
            @current_font[:sz] = attrs["val"]&.to_f if @current_font
          when "color"
            handle_color(attrs)
          when "name"
            @current_font[:name] = attrs["val"] if @current_font
          when "fill"
            @current_fill = {} if parent_context?("fills") || (@current_dxf && parent_context?("dxf"))
          when "patternFill"
            @current_fill_pattern = attrs["patternType"] if @current_fill
          when "fgColor", "bgColor"
            handle_fill_color(localname, attrs) if @current_fill
          when "border"
            if parent_context?("borders") || (@current_dxf && parent_context?("dxf"))
              @current_border = {}
              @current_border[:diagonal_up] = true if attrs["diagonalUp"] == "1"
              @current_border[:diagonal_down] = true if attrs["diagonalDown"] == "1"
            end
          when "left", "right", "top", "bottom", "diagonal"
            handle_border_side(localname, attrs)
          when "cellXfs"
            @in_cell_xfs = true
          when "cellStyleXfs"
            @in_cell_style_xfs = true
          when "xf"
            handle_xf(attrs)
          when "alignment"
            handle_alignment(attrs)
          when "protection"
            handle_protection(attrs)
          end
        end

        def end_element(_uri, localname, _qname)
          case localname
          when "font"
            if @current_font
              if @current_dxf
                @current_dxf[:font] = @current_font
              else
                @result[:fonts] << @current_font
              end
              @current_font = nil
            end
          when "fill"
            if @current_fill
              @current_fill[:pattern] = @current_fill_pattern if @current_fill_pattern
              if @current_dxf
                @current_dxf[:fill] = @current_fill
              else
                @result[:fills] << @current_fill
              end
              @current_fill = nil
              @current_fill_pattern = nil
            end
          when "dxf"
            if @current_dxf
              (@result[:dxfs] ||= []) << @current_dxf
              @current_dxf = nil
            end
          when "border"
            if @current_border
              if @current_dxf
                @current_dxf[:border] = @current_border
              else
                @result[:borders] << @current_border
              end
              @current_border = nil
            end
          when "left", "right", "top", "bottom", "diagonal"
            @current_border_side = nil
          when "cellXfs"
            @in_cell_xfs = false
          when "cellStyleXfs"
            @in_cell_style_xfs = false
          when "xf"
            finalize_xf
          end
          @context.pop
        end

        def characters(_text)
          # No text content needed for styles
        end

        private

        def parent_context?(tag)
          @context.length >= 2 && @context[-2] == tag
        end

        def handle_color(attrs)
          if @current_font
            color = extract_color(attrs)
            @current_font[:color] = color unless color.empty?
          elsif @current_border && @current_border_side
            color = extract_color(attrs)
            @current_border[@current_border_side] ||= {}
            @current_border[@current_border_side][:color] = color unless color.empty?
          end
        end

        def handle_fill_color(localname, attrs)
          color = extract_color(attrs)
          return if color.empty?

          key = localname == "fgColor" ? :fg_color : :bg_color
          @current_fill[key] = color
        end

        def handle_border_side(side, attrs)
          return unless @current_border

          @current_border_side = side.to_sym
          style = attrs["style"]
          @current_border[@current_border_side] ||= {}
          @current_border[@current_border_side][:style] = style if style
        end

        def handle_xf(attrs)
          @current_xf = {
            num_fmt_id: attrs["numFmtId"]&.to_i,
            font_id: attrs["fontId"]&.to_i,
            fill_id: attrs["fillId"]&.to_i,
            border_id: attrs["borderId"]&.to_i,
            xf_id: attrs["xfId"]&.to_i,
            apply_number_format: attrs["applyNumberFormat"] == "1",
            apply_font: attrs["applyFont"] == "1",
            apply_fill: attrs["applyFill"] == "1",
            apply_border: attrs["applyBorder"] == "1",
            apply_alignment: attrs["applyAlignment"] == "1"
          }
        end

        def truthy_val?(attrs)
          val = attrs["val"]
          val.nil? || (val != "0" && val.downcase != "false")
        end

        def handle_alignment(attrs)
          target = @current_xf || (@current_dxf if parent_context?("dxf"))
          return unless target

          alignment = {}
          alignment[:horizontal] = attrs["horizontal"] if attrs["horizontal"]
          alignment[:vertical] = attrs["vertical"] if attrs["vertical"]
          alignment[:wrap_text] = true if attrs["wrapText"] == "1"
          alignment[:text_rotation] = attrs["textRotation"]&.to_i if attrs["textRotation"]
          alignment[:indent] = attrs["indent"]&.to_i if attrs["indent"]
          alignment[:shrink_to_fit] = true if attrs["shrinkToFit"] == "1"
          target[:alignment] = alignment unless alignment.empty?
        end

        def handle_protection(attrs)
          target = @current_xf || (@current_dxf if parent_context?("dxf"))
          return unless target

          protection = {}
          protection[:locked] = attrs["locked"] == "1" unless attrs["locked"].nil?
          protection[:hidden] = attrs["hidden"] == "1" unless attrs["hidden"].nil?
          target[:protection] = protection unless protection.empty?
        end

        def finalize_xf
          return unless @current_xf

          if @in_cell_xfs
            @result[:cell_xfs] << @current_xf
          elsif @in_cell_style_xfs
            @result[:cell_style_xfs] << @current_xf
          end
          @current_xf = nil
        end

        def extract_color(attrs)
          color = {}
          color[:auto] = true if %w[1 true].include?(attrs["auto"])
          color[:rgb] = attrs["rgb"] if attrs["rgb"]
          color[:theme] = attrs["theme"]&.to_i if attrs["theme"]
          color[:indexed] = attrs["indexed"]&.to_i if attrs["indexed"]
          color[:tint] = attrs["tint"]&.to_f if attrs["tint"]
          color
        end
      end
    end
  end
end
