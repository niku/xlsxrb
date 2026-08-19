# frozen_string_literal: true

# rbs_inline: enabled

require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

module Xlsxrb
  module Ooxml
    class Reader
      # SAX2 listener for parsing drawing XML to extract image information.
      class DrawingImagesListener
        include REXML::SAX2Listener

        attr_reader :images

        def initialize
          @images = []
          @current_image = nil
          @inside_anchor = false
          @inside_pic = false
          @inside_from = false
          @inside_to = false
          @current_field = nil
          @text_buffer = +""
          @anchor_from = {}
          @anchor_to = {}
          @inside_ln = false
          @inside_solid_fill = false
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "twoCellAnchor", "oneCellAnchor"
            @inside_anchor = true
            @anchor_from = {}
            @anchor_to = {}
            @anchor_edit_as = attributes["editAs"]
            @anchor_published = %w[1 true].include?(attributes["fPublished"])
          when "pic"
            @inside_pic = true
            @current_image = {}
            @current_image[:macro] = attributes["macro"] if attributes["macro"] && !attributes["macro"].empty?
          when "cNvPr"
            if @inside_pic && @current_image
              @current_image[:name] = attributes["name"] if attributes["name"]
              @current_image[:id] = attributes["id"]&.to_i
              @current_image[:description] = attributes["descr"] if attributes["descr"]
              @current_image[:title] = attributes["title"] if attributes["title"]
              @current_image[:hidden] = %w[1 true].include?(attributes["hidden"]) if attributes["hidden"]
            end
          when "blip"
            rid = attributes["r:embed"] || attributes["embed"]
            @current_image[:embed_rid] = rid if @inside_pic && @current_image && rid
          when "alphaModFix"
            @current_image[:alpha_mod_fix] = attributes["amt"].to_i if @inside_pic && @current_image && attributes["amt"]
          when "srcRect"
            if @inside_pic && @current_image
              sr = {}
              sr[:top] = attributes["t"].to_i if attributes["t"]
              sr[:bottom] = attributes["b"].to_i if attributes["b"]
              sr[:left] = attributes["l"].to_i if attributes["l"]
              sr[:right] = attributes["r"].to_i if attributes["r"]
              @current_image[:src_rect] = sr unless sr.empty?
            end
          when "picLocks"
            if @inside_pic && @current_image
              @current_image[:no_change_aspect] = true if %w[1 true].include?(attributes["noChangeAspect"])
              @current_image[:no_crop] = true if %w[1 true].include?(attributes["noCrop"])
            end
          when "ln"
            if @inside_pic && @current_image
              @inside_ln = true
              @current_image[:line_width] = attributes["w"].to_i if attributes["w"]
            end
          when "solidFill"
            @inside_solid_fill = true if @inside_pic
          when "srgbClr"
            @current_image[:line_color] = attributes["val"] if @inside_pic && @current_image && @inside_solid_fill && @inside_ln && attributes["val"]
          when "schemeClr"
            @current_image[:line_color] = { scheme: attributes["val"] } if @inside_pic && @current_image && @inside_solid_fill && @inside_ln && attributes["val"]
          when "xfrm"
            @current_image[:rotation] = attributes["rot"].to_i if @inside_pic && @current_image && attributes["rot"]
          when "from"
            @inside_from = true if @inside_anchor
          when "to"
            @inside_to = true if @inside_anchor
          when "ext"
            if @inside_pic && @current_image
              cx = attributes["cx"]
              cy = attributes["cy"]
              @current_image[:cx] = cx.to_i if cx
              @current_image[:cy] = cy.to_i if cy
            end
          when "col", "colOff", "row", "rowOff"
            @current_field = name
            @text_buffer = +""
          when "clientData"
            if @inside_anchor
              @anchor_locks_with_sheet = attributes["fLocksWithSheet"]
              @anchor_prints_with_sheet = attributes["fPrintsWithSheet"]
            end
          end
        end

        def characters(text)
          @text_buffer << text if @current_field
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "pic"
            if @current_image && !@current_image.empty?
              @anchor_from.each { |k, v| @current_image[:"from_#{k}"] = v }
              @anchor_to.each { |k, v| @current_image[:"to_#{k}"] = v }
              @current_image[:edit_as] = @anchor_edit_as if @anchor_edit_as
              @images << @current_image
            end
            @current_image = nil
            @inside_pic = false
          when "twoCellAnchor", "oneCellAnchor"
            @images.last[:locks_with_sheet] = @anchor_locks_with_sheet == "1" if @anchor_locks_with_sheet && !@images.empty?
            @images.last[:prints_with_sheet] = @anchor_prints_with_sheet == "1" if @anchor_prints_with_sheet && !@images.empty?
            @images.last[:published] = true if @anchor_published && !@images.empty?
            @inside_anchor = false
            @anchor_from = {}
            @anchor_to = {}
            @anchor_locks_with_sheet = nil
            @anchor_prints_with_sheet = nil
            @anchor_published = false
          when "from"
            @inside_from = false
          when "to"
            @inside_to = false
          when "ln"
            @inside_ln = false
          when "solidFill"
            @inside_solid_fill = false
          when "col", "colOff", "row", "rowOff"
            if @current_field
              val = @text_buffer.to_i
              if @inside_from
                @anchor_from[@current_field] = val
              elsif @inside_to
                @anchor_to[@current_field] = val
              end
            end
            @current_field = nil
          end
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

      # SAX2 listener for parsing drawing XML to extract chart references.
      class DrawingChartsListener
        include REXML::SAX2Listener

        attr_reader :charts

        def initialize
          @charts = []
          @inside_graphic_frame = false
          @current_chart = nil
          @anchor_edit_as = nil
          @inside_anchor = false
          @inside_from = false
          @inside_to = false
          @anchor_from = {}
          @anchor_to = {}
          @current_field = nil
          @text_buffer = +""
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "twoCellAnchor", "oneCellAnchor"
            @anchor_edit_as = attributes["editAs"]
            @anchor_published = %w[1 true].include?(attributes["fPublished"])
            @inside_anchor = true
            @anchor_from = {}
            @anchor_to = {}
          when "graphicFrame"
            @inside_graphic_frame = true
            @current_chart = {}
            @current_chart[:frame_macro] = attributes["macro"] if attributes["macro"] && !attributes["macro"].empty?
          when "cNvPr"
            if @inside_graphic_frame && @current_chart
              @current_chart[:name] = attributes["name"] if attributes["name"]
              @current_chart[:description] = attributes["descr"] if attributes["descr"]
              @current_chart[:frame_title] = attributes["title"] if attributes["title"]
              @current_chart[:frame_hidden] = %w[1 true].include?(attributes["hidden"]) if attributes["hidden"]
            end
          when "chart"
            rid = attributes["r:id"] || attributes["id"]
            @current_chart[:rid] = rid if @inside_graphic_frame && @current_chart && rid
          when "graphicFrameLocks"
            @current_chart[:frame_no_grp] = true if @inside_graphic_frame && @current_chart && %w[1 true].include?(attributes["noGrp"])
          when "from"
            @inside_from = true if @inside_anchor
          when "to"
            @inside_to = true if @inside_anchor
          when "col", "colOff", "row", "rowOff"
            @current_field = name
            @text_buffer = +""
          when "clientData"
            @anchor_locks_with_sheet = attributes["fLocksWithSheet"]
            @anchor_prints_with_sheet = attributes["fPrintsWithSheet"]
          end
        end

        def characters(text)
          @text_buffer << text if @current_field
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "col", "colOff", "row", "rowOff"
            val = @text_buffer.strip.to_i
            @anchor_from[name] = val if @inside_from
            @anchor_to[name] = val if @inside_to
            @current_field = nil
          when "from"
            @inside_from = false
          when "to"
            @inside_to = false
          when "graphicFrame"
            if @current_chart && @current_chart[:rid]
              @current_chart[:edit_as] = @anchor_edit_as if @anchor_edit_as
              @current_chart[:from_col] = @anchor_from["col"] if @anchor_from["col"]
              @current_chart[:from_row] = @anchor_from["row"] if @anchor_from["row"]
              @current_chart[:from_col_off] = @anchor_from["colOff"] if @anchor_from["colOff"]
              @current_chart[:from_row_off] = @anchor_from["rowOff"] if @anchor_from["rowOff"]
              @current_chart[:to_col] = @anchor_to["col"] if @anchor_to["col"]
              @current_chart[:to_row] = @anchor_to["row"] if @anchor_to["row"]
              @current_chart[:to_col_off] = @anchor_to["colOff"] if @anchor_to["colOff"]
              @current_chart[:to_row_off] = @anchor_to["rowOff"] if @anchor_to["rowOff"]
              @charts << @current_chart
            end
            @current_chart = nil
            @inside_graphic_frame = false
          when "twoCellAnchor", "oneCellAnchor"
            @charts.last[:locks_with_sheet] = @anchor_locks_with_sheet == "1" if @anchor_locks_with_sheet && !@charts.empty?
            @charts.last[:prints_with_sheet] = @anchor_prints_with_sheet == "1" if @anchor_prints_with_sheet && !@charts.empty?
            @charts.last[:published] = true if @anchor_published && !@charts.empty?
            @inside_anchor = false
            @anchor_edit_as = nil
            @anchor_locks_with_sheet = nil
            @anchor_prints_with_sheet = nil
            @anchor_published = false
            @anchor_from = {}
            @anchor_to = {}
          end
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

      # SAX2 listener for parsing drawing XML to extract shape elements.
      class DrawingShapesListener
        include REXML::SAX2Listener

        attr_reader :shapes

        def initialize
          @shapes = []
          @inside_anchor = false
          @inside_sp = false
          @current_shape = nil
          @inside_from = false
          @inside_to = false
          @inside_tx_body = false
          @inside_t = false
          @current_field = nil
          @text_buffer = +""
          @anchor_from = {}
          @anchor_to = {}
          @inside_solid_fill = false
          @inside_highlight = false
          @inside_ln = false
          @inside_rpr_ln = false
          @inside_rpr_effect_lst = false
          @inside_cust_dash = false
          @inside_prst_geom = false
          @inside_rpr = false
          @inside_end_para_rpr = false
          @inside_def_rpr = false
          @current_text_font = nil
          @inside_effect_lst = false
          @inside_outer_shdw = false
          @inside_inner_shdw = false
          @inside_glow = false
          @inside_grad_fill = false
          @inside_patt_fill = false
          @inside_fg_clr = false
          @inside_bg_clr = false
          @current_gs_pos = nil
          @inside_spc_bef = false
          @inside_spc_aft = false
          @inside_lnspc = false
          @inside_tab_lst = false
          @inside_bu_clr = false
          @current_paragraph = nil
          @paragraphs_for_shape = nil
          @inside_r = false
          @current_run = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "twoCellAnchor", "oneCellAnchor"
            @inside_anchor = true
            @anchor_from = {}
            @anchor_to = {}
            @anchor_edit_as = attributes["editAs"]
            @anchor_published = %w[1 true].include?(attributes["fPublished"])
          when "sp"
            @inside_sp = true
            @current_shape = {}
            @current_shape[:macro] = attributes["macro"] if attributes["macro"] && !attributes["macro"].empty?
            @current_shape[:textlink] = attributes["textlink"] if attributes["textlink"] && !attributes["textlink"].empty?
          when "cNvPr"
            if @inside_sp && @current_shape
              @current_shape[:name] = attributes["name"] if attributes["name"]
              @current_shape[:id] = attributes["id"]&.to_i
              @current_shape[:description] = attributes["descr"] if attributes["descr"]
              @current_shape[:title] = attributes["title"] if attributes["title"]
              @current_shape[:hidden] = %w[1 true].include?(attributes["hidden"]) if attributes["hidden"]
            end
          when "prstGeom"
            if @inside_sp && @current_shape && attributes["prst"]
              @current_shape[:preset] = attributes["prst"]
              @inside_prst_geom = true
            end
          when "gd"
            if @inside_prst_geom && @inside_sp && @current_shape && attributes["name"] && attributes["fmla"]
              @current_shape[:adjust_values] ||= []
              @current_shape[:adjust_values] << { name: attributes["name"], fmla: attributes["fmla"] }
            end
          when "xfrm"
            @current_shape[:rotation] = attributes["rot"].to_i if @inside_sp && @current_shape && attributes["rot"]
          when "solidFill"
            @inside_solid_fill = true if @inside_sp
          when "highlight"
            @inside_highlight = true if @inside_rpr && @inside_sp
          when "gradFill"
            if @inside_sp && @current_shape && !@inside_ln
              @inside_grad_fill = true
              gf = { stops: [] }
              gf[:rot_with_shape] = %w[1 true].include?(attributes["rotWithShape"]) if attributes["rotWithShape"]
              gf[:flip] = attributes["flip"] if attributes["flip"]
              @current_shape[:gradient_fill] = gf
            end
          when "pattFill"
            if @inside_sp && @current_shape && !@inside_ln
              @inside_patt_fill = true
              @current_shape[:pattern_fill] = { preset: attributes["prst"] } if attributes["prst"]
            end
          when "fgClr"
            @inside_fg_clr = true if @inside_patt_fill
          when "bgClr"
            @inside_bg_clr = true if @inside_patt_fill
          when "gs"
            @current_gs_pos = attributes["pos"].to_i if @inside_grad_fill && attributes["pos"]
          when "lin"
            if @inside_grad_fill && @current_shape
              @current_shape[:gradient_fill][:angle] = attributes["ang"].to_i if attributes["ang"]
              @current_shape[:gradient_fill][:scaled] = %w[1 true].include?(attributes["scaled"]) if attributes["scaled"]
            end
          when "path"
            @current_shape[:gradient_fill][:path] = attributes["path"] if @inside_grad_fill && @current_shape && attributes["path"]
          when "tileRect"
            if @inside_grad_fill && @current_shape
              tr = {}
              %w[l t r b].each { |a| tr[a.to_sym] = attributes[a] if attributes[a] }
              @current_shape[:gradient_fill][:tile_rect] = tr unless tr.empty?
            end
          when "ln"
            if @inside_rpr && @current_text_font
              @inside_ln = true
              @inside_rpr_ln = true
              @current_text_font[:line_width] = attributes["w"].to_i if attributes["w"]
              @current_text_font[:line_cap] = attributes["cap"] if attributes["cap"]
            elsif @inside_sp && @current_shape
              @inside_ln = true
              @current_shape[:line_width] = attributes["w"].to_i if attributes["w"]
              @current_shape[:line_cap] = attributes["cap"] if attributes["cap"]
              @current_shape[:line_align] = attributes["algn"] if attributes["algn"]
              @current_shape[:line_compound] = attributes["cmpd"] if attributes["cmpd"]
            end
          when "effectLst"
            if @inside_rpr && @current_text_font
              @inside_rpr_effect_lst = true
            elsif @inside_sp
              @inside_effect_lst = true
            end
          when "outerShdw"
            if @inside_rpr_effect_lst && @current_text_font
              @inside_outer_shdw = true
              os = {}
              os[:blur_rad] = attributes["blurRad"].to_i if attributes["blurRad"]
              os[:dist] = attributes["dist"].to_i if attributes["dist"]
              os[:dir] = attributes["dir"].to_i if attributes["dir"]
              os[:algn] = attributes["algn"] if attributes["algn"]
              os[:rot_with_shape] = %w[1 true].include?(attributes["rotWithShape"]) if attributes["rotWithShape"]
              @current_text_font[:outer_shadow] = os
            elsif @inside_sp && @inside_effect_lst && @current_shape
              @inside_outer_shdw = true
              os = {}
              os[:blur_rad] = attributes["blurRad"].to_i if attributes["blurRad"]
              os[:dist] = attributes["dist"].to_i if attributes["dist"]
              os[:dir] = attributes["dir"].to_i if attributes["dir"]
              os[:algn] = attributes["algn"] if attributes["algn"]
              os[:rot_with_shape] = %w[1 true].include?(attributes["rotWithShape"]) if attributes["rotWithShape"]
              @current_shape[:outer_shadow] = os
            end
          when "innerShdw"
            if @inside_rpr_effect_lst && @current_text_font
              @inside_inner_shdw = true
              is = {}
              is[:blur_rad] = attributes["blurRad"].to_i if attributes["blurRad"]
              is[:dist] = attributes["dist"].to_i if attributes["dist"]
              is[:dir] = attributes["dir"].to_i if attributes["dir"]
              @current_text_font[:inner_shadow] = is
            elsif @inside_sp && @inside_effect_lst && @current_shape
              @inside_inner_shdw = true
              is = {}
              is[:blur_rad] = attributes["blurRad"].to_i if attributes["blurRad"]
              is[:dist] = attributes["dist"].to_i if attributes["dist"]
              is[:dir] = attributes["dir"].to_i if attributes["dir"]
              @current_shape[:inner_shadow] = is
            end
          when "glow"
            if @inside_rpr_effect_lst && @current_text_font
              @inside_glow = true
              gl = {}
              gl[:rad] = attributes["rad"].to_i if attributes["rad"]
              @current_text_font[:glow] = gl
            elsif @inside_sp && @inside_effect_lst && @current_shape
              @inside_glow = true
              gl = {}
              gl[:rad] = attributes["rad"].to_i if attributes["rad"]
              @current_shape[:glow] = gl
            end
          when "softEdge"
            if @inside_rpr_effect_lst && @current_text_font
              se = {}
              se[:rad] = attributes["rad"].to_i if attributes["rad"]
              @current_text_font[:soft_edge] = se
            elsif @inside_sp && @inside_effect_lst && @current_shape
              se = {}
              se[:rad] = attributes["rad"].to_i if attributes["rad"]
              @current_shape[:soft_edge] = se
            end
          when "blur"
            if @inside_rpr_effect_lst && @current_text_font
              bl = {}
              bl[:rad] = attributes["rad"].to_i if attributes["rad"]
              bl[:grow] = %w[1 true].include?(attributes["grow"]) if attributes["grow"]
              @current_text_font[:blur] = bl
            elsif @inside_sp && @inside_effect_lst && @current_shape
              bl = {}
              bl[:rad] = attributes["rad"].to_i if attributes["rad"]
              bl[:grow] = %w[1 true].include?(attributes["grow"]) if attributes["grow"]
              @current_shape[:blur] = bl
            end
          when "reflection"
            if @inside_rpr_effect_lst && @current_text_font
              rf = {}
              rf[:blur_rad] = attributes["blurRad"].to_i if attributes["blurRad"]
              rf[:st_a] = attributes["stA"].to_i if attributes["stA"]
              rf[:end_a] = attributes["endA"].to_i if attributes["endA"]
              rf[:dist] = attributes["dist"].to_i if attributes["dist"]
              rf[:dir] = attributes["dir"].to_i if attributes["dir"]
              rf[:fade_dir] = attributes["fadeDir"].to_i if attributes["fadeDir"]
              rf[:sx] = attributes["sx"].to_i if attributes["sx"]
              rf[:sy] = attributes["sy"].to_i if attributes["sy"]
              rf[:kx] = attributes["kx"].to_i if attributes["kx"]
              rf[:ky] = attributes["ky"].to_i if attributes["ky"]
              rf[:algn] = attributes["algn"] if attributes["algn"]
              rf[:rot_with_shape] = %w[1 true].include?(attributes["rotWithShape"]) if attributes["rotWithShape"]
              @current_text_font[:reflection] = rf
            elsif @inside_sp && @inside_effect_lst && @current_shape
              rf = {}
              rf[:blur_rad] = attributes["blurRad"].to_i if attributes["blurRad"]
              rf[:st_a] = attributes["stA"].to_i if attributes["stA"]
              rf[:end_a] = attributes["endA"].to_i if attributes["endA"]
              rf[:dist] = attributes["dist"].to_i if attributes["dist"]
              rf[:dir] = attributes["dir"].to_i if attributes["dir"]
              rf[:fade_dir] = attributes["fadeDir"].to_i if attributes["fadeDir"]
              rf[:sx] = attributes["sx"].to_i if attributes["sx"]
              rf[:sy] = attributes["sy"].to_i if attributes["sy"]
              rf[:kx] = attributes["kx"].to_i if attributes["kx"]
              rf[:ky] = attributes["ky"].to_i if attributes["ky"]
              rf[:algn] = attributes["algn"] if attributes["algn"]
              rf[:rot_with_shape] = %w[1 true].include?(attributes["rotWithShape"]) if attributes["rotWithShape"]
              @current_shape[:reflection] = rf
            end
          when "prstDash"
            if @inside_rpr_ln && @current_text_font && attributes["val"]
              @current_text_font[:line_dash] = attributes["val"]
            elsif @inside_sp && @inside_ln && @current_shape && attributes["val"]
              @current_shape[:line_dash] = attributes["val"]
            end
          when "custDash"
            @inside_cust_dash = true if @inside_sp && @inside_ln && @current_shape
          when "ds"
            if @inside_cust_dash && @inside_sp && @inside_ln && @current_shape
              ds = {}
              ds[:d] = attributes["d"].to_i if attributes["d"]
              ds[:sp] = attributes["sp"].to_i if attributes["sp"]
              @current_shape[:line_custom_dash] ||= []
              @current_shape[:line_custom_dash] << ds
            end
          when "round"
            if @inside_rpr_ln && @current_text_font
              @current_text_font[:line_join] = "round"
            elsif @inside_sp && @inside_ln && @current_shape
              @current_shape[:line_join] = "round"
            end
          when "bevel"
            if @inside_rpr_ln && @current_text_font
              @current_text_font[:line_join] = "bevel"
            elsif @inside_sp && @inside_ln && @current_shape
              @current_shape[:line_join] = "bevel"
            end
          when "miter"
            if @inside_rpr_ln && @current_text_font
              @current_text_font[:line_join] = "miter"
              @current_text_font[:line_miter_limit] = attributes["lim"].to_i if attributes["lim"]
            elsif @inside_sp && @inside_ln && @current_shape
              @current_shape[:line_join] = "miter"
              @current_shape[:line_miter_limit] = attributes["lim"].to_i if attributes["lim"]
            end
          when "headEnd"
            if @inside_sp && @inside_ln && @current_shape
              he = {}
              he[:type] = attributes["type"] if attributes["type"]
              he[:w] = attributes["w"] if attributes["w"]
              he[:len] = attributes["len"] if attributes["len"]
              @current_shape[:head_end] = he
            end
          when "tailEnd"
            if @inside_sp && @inside_ln && @current_shape
              te = {}
              te[:type] = attributes["type"] if attributes["type"]
              te[:w] = attributes["w"] if attributes["w"]
              te[:len] = attributes["len"] if attributes["len"]
              @current_shape[:tail_end] = te
            end
          when "spLocks"
            if @inside_sp && @current_shape
              @current_shape[:f_locks_text] = true if %w[1 true].include?(attributes["fLocksText"])
              @current_shape[:no_grp] = true if %w[1 true].include?(attributes["noGrp"])
              @current_shape[:no_rot] = true if %w[1 true].include?(attributes["noRot"])
            end
          when "srgbClr"
            assign_shape_color(attributes["val"]) if attributes["val"]
          when "schemeClr"
            assign_shape_color({ scheme: attributes["val"] }) if attributes["val"]
          when "noFill"
            if @inside_sp && @current_shape
              if @inside_ln
                @current_shape[:no_line] = true
              else
                @current_shape[:no_fill] = true
              end
            end
          when "alpha"
            if @inside_sp && @current_shape && attributes["val"]
              alpha_val = attributes["val"].to_i
              if @inside_outer_shdw && !@inside_rpr_effect_lst
                @current_shape[:outer_shadow][:alpha] = alpha_val if @current_shape[:outer_shadow]
              elsif @inside_inner_shdw && !@inside_rpr_effect_lst
                @current_shape[:inner_shadow][:alpha] = alpha_val if @current_shape[:inner_shadow]
              elsif @inside_glow && !@inside_rpr_effect_lst
                @current_shape[:glow][:alpha] = alpha_val if @current_shape[:glow]
              elsif @inside_ln && @inside_solid_fill && !@inside_rpr
                @current_shape[:line_alpha] = alpha_val
              elsif @inside_solid_fill && !@inside_ln && !@inside_rpr
                @current_shape[:fill_alpha] = alpha_val
              end
            end
          when "tint", "shade", "lumMod", "lumOff", "satMod", "satOff", "hueMod", "hueOff"
            if @inside_sp && @current_shape && @inside_solid_fill && !@inside_rpr && attributes["val"]
              t = { type: name, val: attributes["val"].to_i }
              if @inside_ln
                (@current_shape[:line_color_transforms] ||= []) << t
              else
                (@current_shape[:fill_color_transforms] ||= []) << t
              end
            end
          when "from"
            @inside_from = true if @inside_anchor
          when "to"
            @inside_to = true if @inside_anchor
          when "txBody"
            @inside_tx_body = true if @inside_sp
            @paragraphs_for_shape = [] if @inside_sp
          when "p"
            @current_paragraph = {} if @inside_tx_body && @inside_sp
          when "r"
            if @inside_tx_body && @inside_sp && @current_paragraph
              @inside_r = true
              @current_run = {}
            end
          when "rPr"
            if @inside_tx_body && @inside_sp && @current_shape
              @inside_rpr = true
              tf = {}
              tf[:bold] = true if %w[1 true].include?(attributes["b"])
              tf[:italic] = true if %w[1 true].include?(attributes["i"])
              tf[:no_proof] = true if %w[1 true].include?(attributes["noProof"])
              tf[:normalize_h] = true if %w[1 true].include?(attributes["normalizeH"])
              tf[:kumimoji] = true if %w[1 true].include?(attributes["kumimoji"])
              tf[:strike] = attributes["strike"] if attributes["strike"]
              tf[:underline] = attributes["u"] if attributes["u"]
              tf[:baseline] = attributes["baseline"].to_i if attributes["baseline"]
              tf[:spacing] = attributes["spc"].to_i if attributes["spc"]
              tf[:kern] = attributes["kern"].to_i if attributes["kern"]
              tf[:cap] = attributes["cap"] if attributes["cap"]
              tf[:lang] = attributes["lang"] if attributes["lang"]
              tf[:alt_lang] = attributes["altLang"] if attributes["altLang"]
              tf[:dirty] = true if %w[1 true].include?(attributes["dirty"])
              tf[:smt_clean] = true if %w[1 true].include?(attributes["smtClean"])
              tf[:err] = true if %w[1 true].include?(attributes["err"])
              tf[:bmk] = attributes["bmk"] if attributes["bmk"]
              tf[:size] = attributes["sz"].to_i if attributes["sz"]
              @current_text_font = tf
            end
          when "endParaRPr"
            if @inside_tx_body && @inside_sp && @current_shape
              @inside_rpr = true
              @inside_end_para_rpr = true
              tf = {}
              tf[:bold] = true if %w[1 true].include?(attributes["b"])
              tf[:italic] = true if %w[1 true].include?(attributes["i"])
              tf[:no_proof] = true if %w[1 true].include?(attributes["noProof"])
              tf[:normalize_h] = true if %w[1 true].include?(attributes["normalizeH"])
              tf[:kumimoji] = true if %w[1 true].include?(attributes["kumimoji"])
              tf[:strike] = attributes["strike"] if attributes["strike"]
              tf[:underline] = attributes["u"] if attributes["u"]
              tf[:baseline] = attributes["baseline"].to_i if attributes["baseline"]
              tf[:spacing] = attributes["spc"].to_i if attributes["spc"]
              tf[:kern] = attributes["kern"].to_i if attributes["kern"]
              tf[:cap] = attributes["cap"] if attributes["cap"]
              tf[:lang] = attributes["lang"] if attributes["lang"]
              tf[:alt_lang] = attributes["altLang"] if attributes["altLang"]
              tf[:dirty] = true if %w[1 true].include?(attributes["dirty"])
              tf[:smt_clean] = true if %w[1 true].include?(attributes["smtClean"])
              tf[:err] = true if %w[1 true].include?(attributes["err"])
              tf[:bmk] = attributes["bmk"] if attributes["bmk"]
              tf[:size] = attributes["sz"].to_i if attributes["sz"]
              @current_text_font = tf
            end
          when "latin"
            @current_text_font[:name] = attributes["typeface"] if @inside_rpr && @current_text_font && attributes["typeface"]
          when "ea"
            @current_text_font[:ea_font] = attributes["typeface"] if @inside_rpr && @current_text_font && attributes["typeface"]
          when "cs"
            @current_text_font[:cs_font] = attributes["typeface"] if @inside_rpr && @current_text_font && attributes["typeface"]
          when "sym"
            @current_text_font[:sym_font] = attributes["typeface"] if @inside_rpr && @current_text_font && attributes["typeface"]
          when "uFillTx"
            @current_text_font[:u_fill_tx] = true if @inside_rpr && @current_text_font
          when "uLnTx"
            @current_text_font[:u_ln_tx] = true if @inside_rpr && @current_text_font
          when "pPr"
            if @inside_tx_body && @inside_sp && @current_paragraph
              @current_paragraph[:align] = attributes["algn"] if attributes["algn"]
              @current_paragraph[:font_align] = attributes["fontAlgn"] if attributes["fontAlgn"]
              @current_paragraph[:def_tab_sz] = attributes["defTabSz"].to_i if attributes["defTabSz"]
              @current_paragraph[:rtl] = %w[1 true].include?(attributes["rtl"]) if attributes["rtl"]
              @current_paragraph[:ea_ln_brk] = %w[1 true].include?(attributes["eaLnBrk"]) if attributes["eaLnBrk"]
              @current_paragraph[:latin_ln_brk] = %w[1 true].include?(attributes["latinLnBrk"]) if attributes["latinLnBrk"]
              @current_paragraph[:hanging_punct] = %w[1 true].include?(attributes["hangingPunct"]) if attributes["hangingPunct"]
              @current_paragraph[:level] = attributes["lvl"].to_i if attributes["lvl"]
              ti = {}
              ti[:left] = attributes["marL"].to_i if attributes["marL"]
              ti[:right] = attributes["marR"].to_i if attributes["marR"]
              ti[:indent] = attributes["indent"].to_i if attributes["indent"]
              @current_paragraph[:indent] = ti unless ti.empty?
            end
          when "defRPr"
            if @inside_tx_body && @inside_sp && @current_shape
              @inside_rpr = true
              @inside_def_rpr = true
              tf = {}
              tf[:bold] = true if %w[1 true].include?(attributes["b"])
              tf[:italic] = true if %w[1 true].include?(attributes["i"])
              tf[:no_proof] = true if %w[1 true].include?(attributes["noProof"])
              tf[:normalize_h] = true if %w[1 true].include?(attributes["normalizeH"])
              tf[:kumimoji] = true if %w[1 true].include?(attributes["kumimoji"])
              tf[:strike] = attributes["strike"] if attributes["strike"]
              tf[:underline] = attributes["u"] if attributes["u"]
              tf[:baseline] = attributes["baseline"].to_i if attributes["baseline"]
              tf[:spacing] = attributes["spc"].to_i if attributes["spc"]
              tf[:kern] = attributes["kern"].to_i if attributes["kern"]
              tf[:cap] = attributes["cap"] if attributes["cap"]
              tf[:lang] = attributes["lang"] if attributes["lang"]
              tf[:alt_lang] = attributes["altLang"] if attributes["altLang"]
              tf[:dirty] = true if %w[1 true].include?(attributes["dirty"])
              tf[:smt_clean] = true if %w[1 true].include?(attributes["smtClean"])
              tf[:err] = true if %w[1 true].include?(attributes["err"])
              tf[:bmk] = attributes["bmk"] if attributes["bmk"]
              tf[:size] = attributes["sz"].to_i if attributes["sz"]
              @current_text_font = tf
            end
          when "spcBef"
            @inside_spc_bef = true if @inside_tx_body && @inside_sp
          when "spcAft"
            @inside_spc_aft = true if @inside_tx_body && @inside_sp
          when "lnSpc"
            @inside_lnspc = true if @inside_tx_body && @inside_sp
          when "tabLst"
            @inside_tab_lst = true if @inside_tx_body && @inside_sp
          when "tab"
            if @inside_tab_lst && @inside_tx_body && @inside_sp && @current_paragraph && attributes["pos"]
              tab = { pos: attributes["pos"].to_i }
              tab[:align] = attributes["algn"] if attributes["algn"]
              @current_paragraph[:tab_stops] ||= []
              @current_paragraph[:tab_stops] << tab
            end
          when "buNone"
            if @inside_tx_body && @inside_sp && @current_paragraph
              @current_paragraph[:bullet] ||= {}
              @current_paragraph[:bullet][:type] = "none"
            end
          when "buClr"
            @inside_bu_clr = true if @inside_tx_body && @inside_sp
          when "buFont"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["typeface"]
              @current_paragraph[:bullet] ||= {}
              @current_paragraph[:bullet][:font] = attributes["typeface"]
            end
          when "buSzPts"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["val"]
              @current_paragraph[:bullet] ||= {}
              @current_paragraph[:bullet][:size_pts] = attributes["val"].to_i
            end
          when "buSzPct"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["val"]
              @current_paragraph[:bullet] ||= {}
              @current_paragraph[:bullet][:size_pct] = attributes["val"].to_i
            end
          when "buChar"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["char"]
              @current_paragraph[:bullet] ||= {}
              @current_paragraph[:bullet][:type] = "char"
              @current_paragraph[:bullet][:char] = attributes["char"]
            end
          when "buAutoNum"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["type"]
              @current_paragraph[:bullet] ||= {}
              @current_paragraph[:bullet][:type] = "auto"
              @current_paragraph[:bullet][:auto_type] = attributes["type"]
              @current_paragraph[:bullet][:start_at] = attributes["startAt"].to_i if attributes["startAt"]
            end
          when "spcPts"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["val"]
              @current_paragraph[:spacing] ||= {}
              if @inside_spc_bef
                @current_paragraph[:spacing][:before] = attributes["val"].to_i
              elsif @inside_spc_aft
                @current_paragraph[:spacing][:after] = attributes["val"].to_i
              elsif @inside_lnspc
                @current_paragraph[:spacing][:line] = attributes["val"].to_i
              end
            end
          when "spcPct"
            if @inside_tx_body && @inside_sp && @current_paragraph && attributes["val"]
              @current_paragraph[:spacing] ||= {}
              if @inside_spc_bef
                @current_paragraph[:spacing][:before_pct] = attributes["val"].to_i
              elsif @inside_spc_aft
                @current_paragraph[:spacing][:after_pct] = attributes["val"].to_i
              elsif @inside_lnspc
                @current_paragraph[:spacing][:line_pct] = attributes["val"].to_i
              end
            end
          when "bodyPr"
            if @inside_sp && @current_shape
              @current_shape[:text_rot] = attributes["rot"].to_i if attributes["rot"]
              @current_shape[:text_spc_first_last_para] = %w[1 true].include?(attributes["spcFirstLastPara"]) if attributes["spcFirstLastPara"]
              @current_shape[:text_wrap] = attributes["wrap"] if attributes["wrap"]
              @current_shape[:text_anchor] = attributes["anchor"] if attributes["anchor"]
              @current_shape[:text_anchor_ctr] = %w[1 true].include?(attributes["anchorCtr"]) if attributes["anchorCtr"]
              @current_shape[:text_vert_overflow] = attributes["vertOverflow"] if attributes["vertOverflow"]
              @current_shape[:text_horz_overflow] = attributes["horzOverflow"] if attributes["horzOverflow"]
              @current_shape[:text_num_col] = attributes["numCol"].to_i if attributes["numCol"]
              @current_shape[:text_spc_col] = attributes["spcCol"].to_i if attributes["spcCol"]
              @current_shape[:text_rtl_col] = %w[1 true].include?(attributes["rtlCol"]) if attributes["rtlCol"]
              @current_shape[:text_from_word_art] = %w[1 true].include?(attributes["fromWordArt"]) if attributes["fromWordArt"]
              @current_shape[:text_upright] = %w[1 true].include?(attributes["upright"]) if attributes["upright"]
              @current_shape[:text_compat_ln_spc] = %w[1 true].include?(attributes["compatLnSpc"]) if attributes["compatLnSpc"]
              @current_shape[:text_force_aa] = %w[1 true].include?(attributes["forceAA"]) if attributes["forceAA"]
              @current_shape[:text_vertical] = attributes["vert"] if attributes["vert"]
              ins = {}
              ins[:left] = attributes["lIns"].to_i if attributes["lIns"]
              ins[:top] = attributes["tIns"].to_i if attributes["tIns"]
              ins[:right] = attributes["rIns"].to_i if attributes["rIns"]
              ins[:bottom] = attributes["bIns"].to_i if attributes["bIns"]
              @current_shape[:text_insets] = ins unless ins.empty?
            end
          when "noAutofit"
            @current_shape[:autofit] = "none" if @inside_sp && @current_shape
          when "spAutoFit"
            @current_shape[:autofit] = "shape" if @inside_sp && @current_shape
          when "normAutofit"
            if @inside_sp && @current_shape
              if attributes["fontScale"] || attributes["lnSpcReduction"]
                af = { type: "normal" }
                af[:font_scale] = attributes["fontScale"].to_i if attributes["fontScale"]
                af[:ln_spc_reduction] = attributes["lnSpcReduction"].to_i if attributes["lnSpcReduction"]
                @current_shape[:autofit] = af
              else
                @current_shape[:autofit] = "normal"
              end
            end
          when "prstTxWarp"
            @current_shape[:text_warp] = { preset: attributes["prst"] } if @inside_sp && @current_shape && attributes["prst"]
          when "t"
            @inside_t = true if @inside_tx_body
            @text_buffer = +""
          when "col", "colOff", "row", "rowOff"
            @current_field = name
            @text_buffer = +""
          when "clientData"
            if @inside_anchor
              @anchor_locks_with_sheet = attributes["fLocksWithSheet"]
              @anchor_prints_with_sheet = attributes["fPrintsWithSheet"]
            end
          end
        end

        def characters(text)
          @text_buffer << text if @current_field || @inside_t
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "sp"
            if @current_shape && !@current_shape.empty?
              @anchor_from.each { |k, v| @current_shape[:"from_#{k}"] = v }
              @anchor_to.each { |k, v| @current_shape[:"to_#{k}"] = v }
              @current_shape[:edit_as] = @anchor_edit_as if @anchor_edit_as
              @shapes << @current_shape
            end
            @current_shape = nil
            @inside_sp = false
            @inside_tx_body = false
          when "twoCellAnchor", "oneCellAnchor"
            @shapes.last[:locks_with_sheet] = @anchor_locks_with_sheet == "1" if @anchor_locks_with_sheet && !@shapes.empty?
            @shapes.last[:prints_with_sheet] = @anchor_prints_with_sheet == "1" if @anchor_prints_with_sheet && !@shapes.empty?
            @shapes.last[:published] = true if @anchor_published && !@shapes.empty?
            @inside_anchor = false
            @anchor_from = {}
            @anchor_to = {}
            @anchor_locks_with_sheet = nil
            @anchor_prints_with_sheet = nil
            @anchor_published = false
          when "from"
            @inside_from = false
          when "to"
            @inside_to = false
          when "solidFill"
            @inside_solid_fill = false
          when "highlight"
            @inside_highlight = false
          when "ln"
            @inside_ln = false
            @inside_rpr_ln = false
            @inside_cust_dash = false
          when "effectLst"
            @inside_effect_lst = false
            @inside_rpr_effect_lst = false
          when "outerShdw"
            @inside_outer_shdw = false
          when "innerShdw"
            @inside_inner_shdw = false
          when "glow"
            @inside_glow = false
          when "gradFill"
            @inside_grad_fill = false
          when "pattFill"
            @inside_patt_fill = false
          when "fgClr"
            @inside_fg_clr = false
          when "bgClr"
            @inside_bg_clr = false
          when "prstGeom"
            @inside_prst_geom = false
          when "spcBef"
            @inside_spc_bef = false
          when "spcAft"
            @inside_spc_aft = false
          when "lnSpc"
            @inside_lnspc = false
          when "tabLst"
            @inside_tab_lst = false
          when "buClr"
            @inside_bu_clr = false
          when "rPr"
            if @inside_rpr && !@inside_end_para_rpr && !@inside_def_rpr && @current_text_font&.any?
              if @current_run
                @current_run[:font] = @current_text_font
              elsif @current_paragraph
                @current_paragraph[:font] = @current_text_font
              end
            end
            @inside_rpr = false
            @current_text_font = nil
          when "endParaRPr"
            @current_paragraph[:end_para_rpr] = @current_text_font if @inside_end_para_rpr && @current_text_font&.any? && @current_paragraph
            @inside_rpr = false
            @inside_end_para_rpr = false
            @current_text_font = nil
          when "defRPr"
            @current_paragraph[:def_rpr] = @current_text_font if @inside_def_rpr && @current_text_font&.any? && @current_paragraph
            @inside_rpr = false
            @inside_def_rpr = false
            @current_text_font = nil
          when "t"
            if @inside_t && @inside_tx_body
              if @current_run
                @current_run[:text] = (@current_run[:text] || +"") << @text_buffer
              elsif @current_paragraph
                @current_paragraph[:text] = (@current_paragraph[:text] || +"") << @text_buffer
              end
            end
            @inside_t = false
          when "r"
            if @inside_r && @current_run && @current_paragraph
              @current_paragraph[:runs] ||= []
              @current_paragraph[:runs] << @current_run
              @current_run = nil
              @inside_r = false
            end
          when "p"
            if @inside_tx_body && @inside_sp && @current_paragraph && @paragraphs_for_shape
              finalize_paragraph_runs(@current_paragraph)
              @paragraphs_for_shape << @current_paragraph
              merge_paragraph_to_shape(@current_paragraph, @current_shape) if @current_shape
              @current_paragraph = nil
            end
          when "txBody"
            @current_shape[:text_paragraphs] = @paragraphs_for_shape if @inside_tx_body && @inside_sp && @current_shape && @paragraphs_for_shape && @paragraphs_for_shape.size > 1
            @inside_tx_body = false
            @paragraphs_for_shape = nil
          when "col", "colOff", "row", "rowOff"
            if @current_field
              val = @text_buffer.to_i
              if @inside_from
                @anchor_from[@current_field] = val
              elsif @inside_to
                @anchor_to[@current_field] = val
              end
            end
            @current_field = nil
          end
        end

        private

        def assign_shape_color(color_value)
          if @inside_rpr_ln && @current_text_font && @inside_solid_fill
            @current_text_font[:line_color] = color_value
          elsif @inside_rpr && @current_text_font && @inside_solid_fill
            @current_text_font[:color] = color_value
          elsif @inside_rpr && @current_text_font && @inside_highlight
            @current_text_font[:highlight] = color_value
          elsif @inside_bu_clr && @inside_sp && @current_paragraph
            @current_paragraph[:bullet] ||= {}
            @current_paragraph[:bullet][:color] = color_value
          elsif @inside_outer_shdw && @inside_rpr_effect_lst && @current_text_font
            @current_text_font[:outer_shadow][:color] = color_value
          elsif @inside_inner_shdw && @inside_rpr_effect_lst && @current_text_font
            @current_text_font[:inner_shadow][:color] = color_value
          elsif @inside_glow && @inside_rpr_effect_lst && @current_text_font
            @current_text_font[:glow][:color] = color_value
          elsif @inside_outer_shdw && @current_shape
            @current_shape[:outer_shadow][:color] = color_value
          elsif @inside_inner_shdw && @current_shape
            @current_shape[:inner_shadow][:color] = color_value
          elsif @inside_glow && @current_shape
            @current_shape[:glow][:color] = color_value
          elsif @inside_grad_fill && @current_gs_pos && @current_shape
            @current_shape[:gradient_fill][:stops] << { pos: @current_gs_pos, color: color_value }
            @current_gs_pos = nil
          elsif @inside_patt_fill && @current_shape
            if @inside_fg_clr
              @current_shape[:pattern_fill][:fg_color] = color_value
            elsif @inside_bg_clr
              @current_shape[:pattern_fill][:bg_color] = color_value
            end
          elsif @inside_sp && @current_shape && @inside_solid_fill && !@inside_rpr
            if @inside_ln
              @current_shape[:line_color] = color_value
            else
              @current_shape[:fill_color] = color_value
            end
          end
        end

        def finalize_paragraph_runs(para)
          return unless para[:runs]&.any?

          # Set paragraph text from concatenation of all run texts
          para[:text] = para[:runs].filter_map { |r| r[:text] }.join unless para[:text]
          # Set paragraph font from last run's font for backward compat
          last_font = para[:runs].rfind { |r| r[:font] }&.dig(:font)
          para[:font] = last_font if last_font && !para[:font]
          # Remove :runs key when only one run (simplify output)
          para.delete(:runs) if para[:runs].size <= 1
        end

        def merge_paragraph_to_shape(para, shape)
          shape[:text] = shape[:text] ? "#{shape[:text]}\n#{para[:text]}" : para[:text] if para[:text]
          shape[:text_font] = para[:font] if para[:font]
          shape[:text_align] = para[:align] if para[:align]
          shape[:text_font_align] = para[:font_align] if para[:font_align]
          shape[:text_def_tab_sz] = para[:def_tab_sz] if para[:def_tab_sz]
          shape[:text_rtl] = para[:rtl] unless para[:rtl].nil?
          shape[:text_ea_ln_brk] = para[:ea_ln_brk] unless para[:ea_ln_brk].nil?
          shape[:text_latin_ln_brk] = para[:latin_ln_brk] unless para[:latin_ln_brk].nil?
          shape[:text_hanging_punct] = para[:hanging_punct] unless para[:hanging_punct].nil?
          shape[:text_level] = para[:level] if para[:level]
          shape[:text_indent] = para[:indent] if para[:indent]
          shape[:text_spacing] = para[:spacing] if para[:spacing]
          shape[:text_tab_stops] = para[:tab_stops] if para[:tab_stops]
          shape[:text_bullet] = para[:bullet] if para[:bullet]
          shape[:text_end_para_rpr] = para[:end_para_rpr] if para[:end_para_rpr]
          shape[:text_def_rpr] = para[:def_rpr] if para[:def_rpr]
        end

        def element_name(local_name, qname)
          if local_name.nil? || local_name.empty?
            qname.to_s.split(":").last
          else
            local_name
          end
        end
      end

      # SAX2 listener for parsing chart XML to identify chart type and title.
      class ChartTypeListener
        include REXML::SAX2Listener

        attr_reader :chart_type, :title, :title_overlay, :title_font,
                    :title_fill_color, :title_no_fill, :title_line_color, :title_line_width, :title_line_dash,
                    :series, :legend, :data_labels, :cat_axis_title, :val_axis_title,
                    :grouping, :bar_dir, :vary_colors, :plot_vis_only, :disp_blanks_as, :style, :auto_title_deleted,
                    :rounded_corners, :cat_axis_tick_lbl_pos, :val_axis_tick_lbl_pos,
                    :cat_axis_major_gridlines, :val_axis_major_gridlines,
                    :cat_axis_minor_gridlines, :val_axis_minor_gridlines,
                    :show_d_lbls_over_max, :cat_axis_delete, :val_axis_delete,
                    :cat_axis_orientation, :val_axis_orientation,
                    :gap_width, :overlap, :view_3d,
                    :gap_depth, :bar_shape,
                    :bubble_3d, :bubble_scale, :show_neg_bubbles, :size_represents,
                    :cat_axis_num_fmt, :val_axis_num_fmt,
                    :cat_axis_major_tick_mark, :cat_axis_minor_tick_mark,
                    :val_axis_major_tick_mark, :val_axis_minor_tick_mark,
                    :cat_axis_crosses, :val_axis_crosses,
                    :cat_axis_crosses_at, :val_axis_crosses_at,
                    :cat_axis_tick_lbl_skip, :cat_axis_tick_mark_skip,
                    :cat_axis_lbl_offset, :cat_axis_auto, :cat_axis_lbl_algn,
                    :cat_axis_no_multi_lvl_lbl,
                    :val_axis_cross_between, :val_axis_major_unit, :val_axis_minor_unit,
                    :val_axis_disp_units,
                    :cat_axis_scaling_max, :cat_axis_scaling_min,
                    :val_axis_scaling_max, :val_axis_scaling_min,
                    :cat_axis_log_base, :val_axis_log_base,
                    :first_slice_ang, :hole_size,
                    :smooth, :marker,
                    :drop_lines, :hi_low_lines, :ser_lines,
                    :up_down_bars,
                    :scatter_style, :radar_style,
                    :cat_axis_pos, :val_axis_pos,
                    :wireframe,
                    :band_fmts,
                    :of_pie_type, :split_type, :split_pos, :cust_split, :second_pie_size,
                    :data_table,
                    :plot_area_fill, :plot_area_no_fill, :plot_area_line_color, :plot_area_line_width, :plot_area_line_dash,
                    :plot_area_layout,
                    :cat_axis_label_rotation, :val_axis_label_rotation,
                    :cat_axis_font, :val_axis_font,
                    :cat_axis_fill, :cat_axis_no_fill, :val_axis_fill, :val_axis_no_fill,
                    :cat_axis_line_color, :cat_axis_line_width, :cat_axis_line_dash,
                    :val_axis_line_color, :val_axis_line_width, :val_axis_line_dash,
                    :floor, :side_wall, :back_wall,
                    :legend_font,
                    :cat_axis_type, :cat_axis_base_time_unit,
                    :cat_axis_major_time_unit, :cat_axis_minor_time_unit,
                    :cat_axis_major_unit, :cat_axis_minor_unit,
                    :cat_axis_title_font, :val_axis_title_font,
                    :cat_axis_title_fill, :cat_axis_title_no_fill, :cat_axis_title_line_color, :cat_axis_title_line_width, :cat_axis_title_line_dash,
                    :val_axis_title_fill, :val_axis_title_no_fill, :val_axis_title_line_color, :val_axis_title_line_width, :val_axis_title_line_dash,
                    :title_layout, :cat_axis_title_layout, :val_axis_title_layout,
                    :title_rotation, :cat_axis_title_rotation, :val_axis_title_rotation,
                    :chart_fill, :chart_no_fill, :chart_line_color, :chart_line_width, :chart_line_dash,
                    :protection, :print_settings, :chart_font

        CHART_TYPES = %w[barChart lineChart pieChart areaChart scatterChart doughnutChart radarChart
                         bar3DChart line3DChart pie3DChart area3DChart surfaceChart surface3DChart stockChart bubbleChart
                         ofPieChart].freeze

        def initialize
          @chart_type = nil
          @title = nil
          @title_overlay = nil
          @title_font = nil
          @title_fill_color = nil
          @title_no_fill = nil
          @title_line_color = nil
          @title_line_width = nil
          @title_line_dash = nil
          @title_layout = nil
          @cat_axis_title_layout = nil
          @val_axis_title_layout = nil
          @title_rotation = nil
          @cat_axis_title_rotation = nil
          @val_axis_title_rotation = nil
          @inside_title_layout = false
          @inside_title_sp_pr = false
          @inside_title_ln = false
          @inside_title_solid_fill = false
          @chart_fill = nil
          @chart_no_fill = nil
          @chart_line_color = nil
          @chart_line_width = nil
          @chart_line_dash = nil
          @inside_chart = false
          @inside_chart_space_sp_pr = false
          @inside_chart_space_ln = false
          @inside_chart_space_solid_fill = false
          @chart_font = nil
          @inside_chart_space_tx_pr = false
          @protection = nil
          @inside_protection = false
          @print_settings = nil
          @inside_print_settings = false
          @inside_ps_header_footer = false
          @inside_ps_odd_header = false
          @inside_ps_odd_footer = false
          @inside_ps_even_header = false
          @inside_ps_even_footer = false
          @inside_ps_first_header = false
          @inside_ps_first_footer = false
          @series = []
          @legend = {}
          @data_labels = {}
          @cat_axis_title = nil
          @val_axis_title = nil
          @grouping = nil
          @bar_dir = nil
          @vary_colors = nil
          @plot_vis_only = nil
          @disp_blanks_as = nil
          @style = nil
          @auto_title_deleted = nil
          @rounded_corners = nil
          @cat_axis_tick_lbl_pos = nil
          @val_axis_tick_lbl_pos = nil
          @cat_axis_major_gridlines = false
          @val_axis_major_gridlines = false
          @cat_axis_minor_gridlines = false
          @val_axis_minor_gridlines = false
          @inside_gridlines = false
          @inside_gridlines_sp_pr = false
          @inside_gridlines_ln = false
          @inside_gridlines_solid_fill = false
          @gridlines_target = nil
          @show_d_lbls_over_max = nil
          @cat_axis_delete = nil
          @val_axis_delete = nil
          @cat_axis_orientation = nil
          @val_axis_orientation = nil
          @gap_width = nil
          @overlap = nil
          @gap_depth = nil
          @bar_shape = nil
          @bubble_3d = nil
          @bubble_scale = nil
          @show_neg_bubbles = nil
          @size_represents = nil
          @view_3d = nil
          @cat_axis_num_fmt = nil
          @val_axis_num_fmt = nil
          @cat_axis_major_tick_mark = nil
          @cat_axis_minor_tick_mark = nil
          @val_axis_major_tick_mark = nil
          @val_axis_minor_tick_mark = nil
          @cat_axis_crosses = nil
          @val_axis_crosses = nil
          @cat_axis_crosses_at = nil
          @val_axis_crosses_at = nil
          @cat_axis_tick_lbl_skip = nil
          @cat_axis_tick_mark_skip = nil
          @cat_axis_lbl_offset = nil
          @cat_axis_auto = nil
          @cat_axis_lbl_algn = nil
          @cat_axis_no_multi_lvl_lbl = nil
          @val_axis_cross_between = nil
          @val_axis_major_unit = nil
          @val_axis_minor_unit = nil
          @val_axis_disp_units = nil
          @inside_disp_units = false
          @inside_disp_units_lbl = false
          @inside_disp_units_lbl_sp_pr = false
          @inside_disp_units_lbl_ln = false
          @inside_disp_units_lbl_solid_fill = false
          @disp_units_lbl_font = nil
          @cat_axis_scaling_max = nil
          @cat_axis_scaling_min = nil
          @val_axis_scaling_max = nil
          @val_axis_scaling_min = nil
          @cat_axis_log_base = nil
          @val_axis_log_base = nil
          @first_slice_ang = nil
          @hole_size = nil
          @smooth = nil
          @marker = nil
          @drop_lines = nil
          @inside_drop_lines = false
          @inside_drop_lines_sp_pr = false
          @inside_drop_lines_ln = false
          @inside_drop_lines_solid_fill = false
          @hi_low_lines = nil
          @inside_hi_low_lines = false
          @inside_hi_low_lines_sp_pr = false
          @inside_hi_low_lines_ln = false
          @inside_hi_low_lines_solid_fill = false
          @ser_lines = nil
          @inside_ser_lines = false
          @inside_ser_lines_sp_pr = false
          @inside_ser_lines_ln = false
          @inside_ser_lines_solid_fill = false
          @up_down_bars = nil
          @inside_up_down_bars = false
          @inside_up_bars = false
          @inside_down_bars = false
          @inside_up_down_bar_sp_pr = false
          @inside_up_down_bar_ln = false
          @inside_up_down_bar_solid_fill = false
          @scatter_style = nil
          @radar_style = nil
          @cat_axis_pos = nil
          @cat_axis_type = nil
          @cat_axis_base_time_unit = nil
          @cat_axis_major_time_unit = nil
          @cat_axis_minor_time_unit = nil
          @cat_axis_major_unit = nil
          @cat_axis_minor_unit = nil
          @val_axis_pos = nil
          @wireframe = nil
          @band_fmts = nil
          @of_pie_type = nil
          @split_type = nil
          @split_pos = nil
          @cust_split = nil
          @inside_cust_split = false
          @second_pie_size = nil
          @inside_band_fmts = false
          @inside_band_fmt = false
          @inside_band_fmt_sp_pr = false
          @inside_band_fmt_ln = false
          @inside_band_fmt_solid_fill = false
          @current_band_fmt = nil
          @data_table = nil
          @plot_area_fill = nil
          @plot_area_no_fill = nil
          @plot_area_line_color = nil
          @plot_area_line_width = nil
          @plot_area_line_dash = nil
          @inside_plot_area = false
          @inside_plot_area_layout = false
          @inside_plot_area_sp_pr = false
          @inside_plot_area_ln = false
          @inside_plot_area_solid_fill = false
          @inside_ax_sp_pr = false
          @inside_ax_ln = false
          @inside_ax_solid_fill = false
          @cat_axis_fill = nil
          @cat_axis_no_fill = nil
          @val_axis_fill = nil
          @val_axis_no_fill = nil
          @cat_axis_line_color = nil
          @cat_axis_line_width = nil
          @cat_axis_line_dash = nil
          @val_axis_line_color = nil
          @val_axis_line_width = nil
          @val_axis_line_dash = nil
          @inside_d_table = false
          @inside_d_table_sp_pr = false
          @inside_d_table_ln = false
          @inside_d_table_solid_fill = false
          @d_table_font = nil
          @inside_view_3d = false
          @inside_wall = false
          @current_wall = nil
          @inside_wall_sp_pr = false
          @inside_wall_ln = false
          @inside_wall_solid_fill = false
          @floor = nil
          @side_wall = nil
          @back_wall = nil
          @inside_scaling = false
          @inside_title = false
          @inside_t = false
          @text_buffer = +""
          @inside_ser = false
          @inside_ser_sp_pr = false
          @inside_ser_solid_fill = false
          @inside_ser_ln = false
          @inside_ser_marker = false
          @inside_marker_sp_pr = false
          @inside_marker_ln = false
          @inside_marker_solid_fill = false
          @inside_dpt = false
          @inside_dpt_marker = false
          @inside_dpt_marker_sp_pr = false
          @inside_dpt_marker_solid_fill = false
          @inside_dpt_marker_ln = false
          @inside_dpt_sp_pr = false
          @inside_dpt_solid_fill = false
          @inside_dpt_ln = false
          @current_dpt = nil
          @inside_trendline = false
          @inside_trendline_name = false
          @current_trendline = nil
          @inside_trendline_sp_pr = false
          @inside_trendline_ln = false
          @inside_trendline_solid_fill = false
          @inside_trendline_lbl = false
          @inside_trendline_lbl_layout = false
          @inside_trendline_lbl_tx = false
          @inside_trendline_lbl_sp_pr = false
          @inside_trendline_lbl_ln = false
          @inside_trendline_lbl_solid_fill = false
          @trendline_lbl_font = nil
          @inside_err_bars = false
          @current_err_bars = nil
          @inside_err_bars_sp_pr = false
          @inside_err_bars_ln = false
          @inside_err_bars_solid_fill = false
          @inside_err_bars_plus = false
          @inside_err_bars_minus = false
          @current_ser = nil
          @inside_cat = false
          @inside_val = false
          @inside_bubble_size = false
          @inside_f = false
          @inside_num_cache = false
          @inside_str_cache = false
          @inside_cache_pt = false
          @inside_cache_v = false
          @current_cache_idx = 0
          @cache_values = []
          @inside_legend = false
          @inside_legend_entry = false
          @inside_legend_layout = false
          @current_legend_entry = nil
          @legend_font = nil
          @inside_legend_sp_pr = false
          @inside_legend_ln = false
          @inside_legend_solid_fill = false
          @inside_dlbls = false
          @inside_dlbl = false
          @inside_dlbl_tx = false
          @inside_dlbl_sp_pr = false
          @inside_dlbl_solid_fill = false
          @inside_dlbl_ln = false
          @inside_dlbl_layout = false
          @current_dlbl = nil
          @inside_dlbls_sp_pr = false
          @inside_dlbls_ln = false
          @inside_dlbls_solid_fill = false
          @dlbls_font = nil
          @dlbl_target = nil
          @inside_separator = false
          @inside_leader_lines = false
          @inside_leader_lines_sp_pr = false
          @inside_leader_lines_ln = false
          @inside_leader_lines_solid_fill = false
          @inside_cat_ax = false
          @inside_val_ax = false
          @inside_ax_title = false
          @inside_ax_title_rpr = false
          @inside_ax_title_sp_pr = false
          @inside_ax_title_ln = false
          @inside_ax_title_solid_fill = false
          @cat_axis_title_font = nil
          @val_axis_title_font = nil
          @cat_axis_title_fill = nil
          @cat_axis_title_no_fill = nil
          @cat_axis_title_line_color = nil
          @cat_axis_title_line_width = nil
          @cat_axis_title_line_dash = nil
          @val_axis_title_fill = nil
          @val_axis_title_no_fill = nil
          @val_axis_title_line_color = nil
          @val_axis_title_line_width = nil
          @val_axis_title_line_dash = nil
          @inside_title_rpr = false
          @title_depth = 0
          @inside_axis_tx_pr = false
          @inside_axis_def_rpr = false
          @cat_axis_font = nil
          @val_axis_font = nil
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          @chart_type = name if CHART_TYPES.include?(name)

          case name
          when "chart"
            @inside_chart = true
          when "plotArea"
            @inside_plot_area = true
          when "grouping"
            @grouping = attributes["val"] if attributes["val"]
          when "barDir"
            @bar_dir = attributes["val"] if attributes["val"]
          when "varyColors"
            @vary_colors = attributes["val"] == "1" if attributes["val"]
          when "autoTitleDeleted"
            @auto_title_deleted = attributes["val"] == "1" if attributes["val"]
          when "view3D"
            @inside_view_3d = true
            @view_3d = {}
          when "rotX"
            @view_3d[:rot_x] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
          when "hPercent"
            @view_3d[:h_percent] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
          when "rotY"
            @view_3d[:rot_y] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
          when "depthPercent"
            @view_3d[:depth_percent] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
          when "rAngAx"
            @view_3d[:r_ang_ax] = attributes["val"] == "1" if @inside_view_3d && attributes["val"]
          when "perspective"
            @view_3d[:perspective] = attributes["val"].to_i if @inside_view_3d && attributes["val"]
          when "floor"
            @inside_wall = true
            @current_wall = (@floor ||= {})
          when "sideWall"
            @inside_wall = true
            @current_wall = (@side_wall ||= {})
          when "backWall"
            @inside_wall = true
            @current_wall = (@back_wall ||= {})
          when "gapWidth"
            if @inside_up_down_bars && attributes["val"]
              @up_down_bars[:gap_width] = attributes["val"].to_i
            elsif attributes["val"]
              @gap_width = attributes["val"].to_i
            end
          when "overlap"
            @overlap = attributes["val"]&.to_i if attributes["val"]
          when "gapDepth"
            @gap_depth = attributes["val"]&.to_i if attributes["val"]
          when "shape"
            if @inside_ser && @current_ser && attributes["val"]
              @current_ser[:shape] = attributes["val"]
            elsif attributes["val"]
              @bar_shape = attributes["val"]
            end
          when "bubble3D"
            if @inside_dpt && @current_dpt && attributes["val"]
              @current_dpt[:bubble_3d] = attributes["val"] == "1"
            elsif attributes["val"] && !@inside_ser
              @bubble_3d = attributes["val"] == "1"
            end
          when "bubbleScale"
            @bubble_scale = attributes["val"]&.to_i if attributes["val"]
          when "showNegBubbles"
            @show_neg_bubbles = attributes["val"] == "1" if attributes["val"]
          when "sizeRepresents"
            @size_represents = attributes["val"] if attributes["val"]
          when "firstSliceAng"
            @first_slice_ang = attributes["val"]&.to_i if attributes["val"]
          when "holeSize"
            @hole_size = attributes["val"]&.to_i if attributes["val"]
          when "smooth"
            if @inside_ser && @current_ser && attributes["val"]
              @current_ser[:smooth] = attributes["val"] == "1"
            elsif attributes["val"]
              @smooth = attributes["val"] == "1"
            end
          when "invertIfNegative"
            if @inside_dpt && @current_dpt && attributes["val"]
              @current_dpt[:invert_if_negative] = attributes["val"] == "1"
            elsif @inside_ser && @current_ser && attributes["val"]
              @current_ser[:invert_if_negative] = attributes["val"] == "1"
            end
          when "explosion"
            if @inside_dpt && @current_dpt && attributes["val"]
              @current_dpt[:explosion] = attributes["val"].to_i
            elsif @inside_ser && @current_ser && attributes["val"]
              @current_ser[:explosion] = attributes["val"].to_i
            end
          when "marker"
            if @inside_dpt
              @inside_dpt_marker = true
            elsif attributes["val"] && !@inside_ser
              @marker = attributes["val"] == "1"
            elsif @inside_ser
              @inside_ser_marker = true
            end
          when "symbol"
            if @inside_dpt_marker && @current_dpt && attributes["val"]
              @current_dpt[:marker_symbol] = attributes["val"]
            elsif @inside_ser_marker && @current_ser && attributes["val"]
              @current_ser[:marker_symbol] = attributes["val"]
            end
          when "size"
            if @inside_dpt_marker && @current_dpt && attributes["val"]
              @current_dpt[:marker_size] = attributes["val"].to_i
            elsif @inside_ser_marker && @current_ser && attributes["val"]
              @current_ser[:marker_size] = attributes["val"].to_i
            end
          when "scatterStyle"
            @scatter_style = attributes["val"] if attributes["val"]
          when "radarStyle"
            @radar_style = attributes["val"] if attributes["val"]
          when "ofPieType"
            @of_pie_type = attributes["val"] if attributes["val"]
          when "splitType"
            @split_type = attributes["val"] if attributes["val"]
          when "splitPos"
            @split_pos = attributes["val"].to_f if attributes["val"]
          when "custSplit"
            @cust_split = []
            @inside_cust_split = true
          when "secondPiePt"
            @cust_split << attributes["val"].to_i if @inside_cust_split && attributes["val"]
          when "secondPieSize"
            @second_pie_size = attributes["val"].to_i if attributes["val"]
          when "dropLines"
            @drop_lines = true
            @inside_drop_lines = true
          when "hiLowLines"
            @hi_low_lines = true
            @inside_hi_low_lines = true
          when "serLines"
            @ser_lines = true
            @inside_ser_lines = true
          when "upDownBars"
            @inside_up_down_bars = true
            @up_down_bars = {}
          when "upBars"
            @inside_up_bars = true if @inside_up_down_bars
          when "downBars"
            @inside_down_bars = true if @inside_up_down_bars
          when "wireframe"
            @wireframe = attributes["val"] == "1" if attributes["val"]
          when "bandFmts"
            @inside_band_fmts = true
            @band_fmts = []
          when "bandFmt"
            if @inside_band_fmts
              @inside_band_fmt = true
              @current_band_fmt = {}
            end
          when "ser"
            @inside_ser = true
            @current_ser = {}
          when "dPt"
            if @inside_ser
              @inside_dpt = true
              @current_dpt = {}
            end
          when "trendline"
            if @inside_ser
              @inside_trendline = true
              @current_trendline = {}
            end
          when "trendlineType"
            @current_trendline[:type] = attributes["val"] if @inside_trendline && @current_trendline && attributes["val"]
          when "order"
            if @inside_trendline && @current_trendline && attributes["val"]
              @current_trendline[:order] = attributes["val"].to_i
            elsif @inside_ser && @current_ser && attributes["val"]
              @current_ser[:order] = attributes["val"].to_i
            end
          when "period"
            @current_trendline[:period] = attributes["val"].to_i if @inside_trendline && @current_trendline && attributes["val"]
          when "forward"
            @current_trendline[:forward] = attributes["val"].to_f if @inside_trendline && @current_trendline && attributes["val"]
          when "backward"
            @current_trendline[:backward] = attributes["val"].to_f if @inside_trendline && @current_trendline && attributes["val"]
          when "intercept"
            @current_trendline[:intercept] = attributes["val"].to_f if @inside_trendline && @current_trendline && attributes["val"]
          when "dispRSqr"
            @current_trendline[:disp_r_sqr] = attributes["val"] == "1" if @inside_trendline && @current_trendline && attributes["val"]
          when "dispEq"
            @current_trendline[:disp_eq] = attributes["val"] == "1" if @inside_trendline && @current_trendline && attributes["val"]
          when "trendlineLbl"
            @inside_trendline_lbl = true if @inside_trendline
          when "errBars"
            if @inside_ser
              @inside_err_bars = true
              @current_err_bars = {}
            end
          when "errDir"
            @current_err_bars[:direction] = attributes["val"] if @inside_err_bars && @current_err_bars && attributes["val"]
          when "errBarType"
            @current_err_bars[:bar_type] = attributes["val"] if @inside_err_bars && @current_err_bars && attributes["val"]
          when "errValType"
            @current_err_bars[:val_type] = attributes["val"] if @inside_err_bars && @current_err_bars && attributes["val"]
          when "noEndCap"
            @current_err_bars[:no_end_cap] = attributes["val"] == "1" if @inside_err_bars && @current_err_bars && attributes["val"]
          when "plus"
            @inside_err_bars_plus = true if @inside_err_bars
          when "minus"
            @inside_err_bars_minus = true if @inside_err_bars
          when "name"
            if @inside_trendline
              @inside_trendline_name = true
              @text_buffer = +""
            end
          when "spPr"
            if @inside_band_fmt
              @inside_band_fmt_sp_pr = true
            elsif @inside_dpt_marker
              @inside_dpt_marker_sp_pr = true
            elsif @inside_dpt
              @inside_dpt_sp_pr = true
            elsif @inside_ser_marker
              @inside_marker_sp_pr = true
            elsif @inside_leader_lines
              @inside_leader_lines_sp_pr = true
            elsif @inside_dlbl
              @inside_dlbl_sp_pr = true
            elsif @inside_dlbls && !@inside_dlbl
              @inside_dlbls_sp_pr = true
            elsif @inside_trendline_lbl
              @inside_trendline_lbl_sp_pr = true
            elsif @inside_trendline
              @inside_trendline_sp_pr = true
            elsif @inside_err_bars
              @inside_err_bars_sp_pr = true
            elsif @inside_ser
              @inside_ser_sp_pr = true
            elsif @inside_drop_lines
              @inside_drop_lines_sp_pr = true
            elsif @inside_hi_low_lines
              @inside_hi_low_lines_sp_pr = true
            elsif @inside_ser_lines
              @inside_ser_lines_sp_pr = true
            elsif @inside_up_bars || @inside_down_bars
              @inside_up_down_bar_sp_pr = true
            elsif @inside_gridlines
              @inside_gridlines_sp_pr = true
            elsif @inside_ax_title
              @inside_ax_title_sp_pr = true
            elsif @inside_disp_units_lbl
              @inside_disp_units_lbl_sp_pr = true
            elsif @inside_cat_ax || @inside_val_ax
              @inside_ax_sp_pr = true
            elsif @inside_wall && @current_wall
              @inside_wall_sp_pr = true
            elsif @inside_legend
              @inside_legend_sp_pr = true
            elsif @inside_d_table
              @inside_d_table_sp_pr = true
            elsif @inside_title && @title_depth == 1 && !@inside_ax_title
              @inside_title_sp_pr = true
            elsif @inside_plot_area
              @inside_plot_area_sp_pr = true
            elsif !@inside_chart
              @inside_chart_space_sp_pr = true
            end
          when "ln"
            if @inside_band_fmt_sp_pr && @current_band_fmt
              @inside_band_fmt_ln = true
              @current_band_fmt[:line_width] = attributes["w"].to_i / 12_700.0 if attributes["w"]
            elsif @inside_dpt_marker_sp_pr && @current_dpt
              @inside_dpt_marker_ln = true
              @current_dpt[:marker_line_width] = attributes["w"].to_i / 12_700.0 if attributes["w"]
            elsif @inside_dpt && @inside_dpt_sp_pr
              @inside_dpt_ln = true
              @current_dpt[:line_width] = attributes["w"].to_i / 12_700.0 if @current_dpt && attributes["w"]
            elsif @inside_marker_sp_pr && @current_ser
              @inside_marker_ln = true
              @current_ser[:marker_line_width] = attributes["w"].to_i / 12_700.0 if attributes["w"]
            elsif @inside_dlbl_sp_pr && @current_dlbl
              @inside_dlbl_ln = true
              @current_dlbl[:line_width] = attributes["w"].to_i / 12_700.0 if attributes["w"]
            elsif @inside_dlbls_sp_pr
              @inside_dlbls_ln = true
              if attributes["w"]
                lw = attributes["w"].to_i / 12_700.0
                dl_target = @dlbl_target || @data_labels
                dl_target[:line_width] = lw
                @data_labels[:line_width] = lw if dl_target != @data_labels
              end
            elsif @inside_trendline_lbl_sp_pr
              @inside_trendline_lbl_ln = true
              (@current_trendline[:label] ||= {})[:line_width] = attributes["w"].to_i / 12_700.0 if @current_trendline && attributes["w"]
            elsif @inside_trendline_sp_pr
              @inside_trendline_ln = true
              @current_trendline[:line_width] = attributes["w"].to_i / 12_700.0 if @current_trendline && attributes["w"]
            elsif @inside_err_bars_sp_pr
              @inside_err_bars_ln = true
              @current_err_bars[:line_width] = attributes["w"].to_i / 12_700.0 if @current_err_bars && attributes["w"]
            elsif @inside_leader_lines_sp_pr
              @inside_leader_lines_ln = true
              if attributes["w"] && @inside_dlbls
                ll = (@dlbl_target[:leader_lines] ||= {})
                ll[:line_width] = attributes["w"].to_i / 12_700.0
                if @dlbl_target != @data_labels
                  ll2 = (@data_labels[:leader_lines] ||= {})
                  ll2[:line_width] = attributes["w"].to_i / 12_700.0
                end
              end
            elsif @inside_ser && @inside_ser_sp_pr
              @inside_ser_ln = true
              @current_ser[:line_width] = attributes["w"].to_i / 12_700.0 if @current_ser && attributes["w"]
              @current_ser[:line_cap] = attributes["cap"] if @current_ser && attributes["cap"]
            elsif @inside_drop_lines_sp_pr
              @inside_drop_lines_ln = true
              if attributes["w"]
                @drop_lines = {} if @drop_lines == true
                @drop_lines[:line_width] = attributes["w"].to_i / 12_700.0
              end
            elsif @inside_hi_low_lines_sp_pr
              @inside_hi_low_lines_ln = true
              if attributes["w"]
                @hi_low_lines = {} if @hi_low_lines == true
                @hi_low_lines[:line_width] = attributes["w"].to_i / 12_700.0
              end
            elsif @inside_ser_lines_sp_pr
              @inside_ser_lines_ln = true
              if attributes["w"]
                @ser_lines = {} if @ser_lines == true
                @ser_lines[:line_width] = attributes["w"].to_i / 12_700.0
              end
            elsif @inside_up_down_bar_sp_pr
              @inside_up_down_bar_ln = true
              bar_key = @inside_up_bars ? :up_bars : :down_bars
              if attributes["w"]
                @up_down_bars[bar_key] ||= {}
                @up_down_bars[bar_key][:line_width] = attributes["w"].to_i / 12_700.0
              end
            elsif @inside_ax_title_sp_pr
              @inside_ax_title_ln = true
              if attributes["w"]
                lw = attributes["w"].to_i / 12_700.0
                if @inside_cat_ax
                  @cat_axis_title_line_width = lw
                elsif @inside_val_ax
                  @val_axis_title_line_width = lw
                end
              end
            elsif @inside_disp_units_lbl_sp_pr
              @inside_disp_units_lbl_ln = true
              if attributes["w"]
                du = @val_axis_disp_units
                du = {} unless du.is_a?(Hash)
                @val_axis_disp_units = du
                (du[:label] ||= {})[:line_width] = attributes["w"].to_i / 12_700.0
              end
            elsif @inside_gridlines_sp_pr
              @inside_gridlines_ln = true
              if attributes["w"] && @gridlines_target
                gl = instance_variable_get(:"@#{@gridlines_target}")
                gl = {} if gl == true
                gl[:line_width] = attributes["w"].to_i / 12_700.0
                instance_variable_set(:"@#{@gridlines_target}", gl)
              end
            elsif @inside_plot_area_sp_pr
              @inside_plot_area_ln = true
              @plot_area_line_width = attributes["w"].to_i if attributes["w"]
            elsif @inside_ax_sp_pr
              @inside_ax_ln = true
              if attributes["w"]
                if @inside_cat_ax
                  @cat_axis_line_width = attributes["w"].to_i
                elsif @inside_val_ax
                  @val_axis_line_width = attributes["w"].to_i
                end
              end
            elsif @inside_wall_sp_pr
              @inside_wall_ln = true
              @current_wall[:line_width] = attributes["w"].to_i if @current_wall && attributes["w"]
            elsif @inside_legend_sp_pr
              @inside_legend_ln = true
              @legend[:line_width] = attributes["w"].to_i / 12_700.0 if attributes["w"]
            elsif @inside_d_table_sp_pr
              @inside_d_table_ln = true
              @data_table[:line_width] = attributes["w"].to_i / 12_700.0 if @data_table && attributes["w"]
            elsif @inside_title_sp_pr
              @inside_title_ln = true
              @title_line_width = attributes["w"].to_i / 12_700.0 if attributes["w"]
            elsif @inside_chart_space_sp_pr
              @inside_chart_space_ln = true
              @chart_line_width = attributes["w"].to_i / 12_700.0 if attributes["w"]
            end
          when "round"
            @current_ser[:line_join] = "round" if @inside_ser && @inside_ser_ln && @current_ser
          when "bevel"
            @current_ser[:line_join] = "bevel" if @inside_ser && @inside_ser_ln && @current_ser
          when "prstDash"
            if @inside_band_fmt_ln && @current_band_fmt && attributes["val"]
              @current_band_fmt[:line_dash] = attributes["val"]
            elsif @inside_gridlines_ln && @gridlines_target && attributes["val"]
              gl = instance_variable_get(:"@#{@gridlines_target}")
              gl = {} if gl == true
              gl[:line_dash] = attributes["val"]
              instance_variable_set(:"@#{@gridlines_target}", gl)
            elsif @inside_trendline_lbl_ln && @current_trendline && attributes["val"]
              (@current_trendline[:label] ||= {})[:line_dash] = attributes["val"]
            elsif @inside_trendline_ln && @current_trendline && attributes["val"]
              @current_trendline[:line_dash] = attributes["val"]
            elsif @inside_err_bars_ln && @current_err_bars && attributes["val"]
              @current_err_bars[:line_dash] = attributes["val"]
            elsif @inside_ser && @inside_ser_ln && @current_ser && attributes["val"]
              @current_ser[:line_dash] = attributes["val"]
            elsif @inside_marker_sp_pr && @inside_marker_ln && @current_ser && attributes["val"]
              @current_ser[:marker_line_dash] = attributes["val"]
            elsif @inside_dpt_marker_sp_pr && @inside_dpt_marker_ln && @current_dpt && attributes["val"]
              @current_dpt[:marker_line_dash] = attributes["val"]
            elsif @inside_drop_lines_ln && attributes["val"]
              @drop_lines = {} if @drop_lines == true
              @drop_lines[:line_dash] = attributes["val"]
            elsif @inside_hi_low_lines_ln && attributes["val"]
              @hi_low_lines = {} if @hi_low_lines == true
              @hi_low_lines[:line_dash] = attributes["val"]
            elsif @inside_ser_lines_ln && attributes["val"]
              @ser_lines = {} if @ser_lines == true
              @ser_lines[:line_dash] = attributes["val"]
            elsif @inside_leader_lines_ln && @inside_dlbls && attributes["val"]
              ll = (@dlbl_target[:leader_lines] ||= {})
              ll[:line_dash] = attributes["val"]
              if @dlbl_target != @data_labels
                ll2 = (@data_labels[:leader_lines] ||= {})
                ll2[:line_dash] = attributes["val"]
              end
            elsif @inside_legend_sp_pr && @inside_legend_ln && attributes["val"]
              @legend[:line_dash] = attributes["val"]
            elsif @inside_d_table_sp_pr && @inside_d_table_ln && attributes["val"]
              @data_table[:line_dash] = attributes["val"]
            elsif @inside_dlbl_sp_pr && @inside_dlbl_ln && @current_dlbl && attributes["val"]
              @current_dlbl[:line_dash] = attributes["val"]
            elsif @inside_dlbls_sp_pr && @inside_dlbls_ln && attributes["val"]
              @data_labels[:line_dash] = attributes["val"]
            elsif @inside_dpt && @inside_dpt_ln && @current_dpt && attributes["val"]
              @current_dpt[:line_dash] = attributes["val"]
            elsif @inside_up_down_bar_sp_pr && @inside_up_down_bar_ln && attributes["val"]
              bar_key = @inside_up_bars ? :up_bars : :down_bars
              @up_down_bars[bar_key] ||= {}
              @up_down_bars[bar_key][:line_dash] = attributes["val"]
            elsif @inside_wall_sp_pr && @inside_wall_ln && @current_wall && attributes["val"]
              @current_wall[:line_dash] = attributes["val"]
            elsif @inside_plot_area_sp_pr && @inside_plot_area_ln && attributes["val"]
              @plot_area_line_dash = attributes["val"]
            elsif @inside_ax_title_sp_pr && @inside_ax_title_ln && attributes["val"]
              if @inside_cat_ax
                @cat_axis_title_line_dash = attributes["val"]
              elsif @inside_val_ax
                @val_axis_title_line_dash = attributes["val"]
              end
            elsif @inside_disp_units_lbl_ln && attributes["val"]
              du = @val_axis_disp_units
              du = {} unless du.is_a?(Hash)
              @val_axis_disp_units = du
              (du[:label] ||= {})[:line_dash] = attributes["val"]
            elsif @inside_ax_sp_pr && @inside_ax_ln && attributes["val"]
              if @inside_cat_ax
                @cat_axis_line_dash = attributes["val"]
              elsif @inside_val_ax
                @val_axis_line_dash = attributes["val"]
              end
            elsif @inside_title_sp_pr && @inside_title_ln && attributes["val"]
              @title_line_dash = attributes["val"]
            elsif @inside_chart_space_sp_pr && @inside_chart_space_ln && attributes["val"]
              @chart_line_dash = attributes["val"]
            end
          when "miter"
            if @inside_ser && @inside_ser_ln && @current_ser
              @current_ser[:line_join] = "miter"
              @current_ser[:line_miter_limit] = attributes["lim"].to_i if attributes["lim"]
            end
          when "solidFill"
            if @inside_band_fmt_sp_pr
              @inside_band_fmt_solid_fill = true
            elsif @inside_dpt_marker_sp_pr
              @inside_dpt_marker_solid_fill = true
            elsif @inside_dpt && @inside_dpt_sp_pr
              @inside_dpt_solid_fill = true
            elsif @inside_marker_sp_pr
              @inside_marker_solid_fill = true
            elsif @inside_dlbl_sp_pr
              @inside_dlbl_solid_fill = true
            elsif @inside_dlbls_sp_pr
              @inside_dlbls_solid_fill = true
            elsif @inside_trendline_lbl_sp_pr
              @inside_trendline_lbl_solid_fill = true
            elsif @inside_trendline_sp_pr
              @inside_trendline_solid_fill = true
            elsif @inside_err_bars_sp_pr
              @inside_err_bars_solid_fill = true
            elsif @inside_leader_lines_sp_pr
              @inside_leader_lines_solid_fill = true
            elsif @inside_ser && @inside_ser_sp_pr
              @inside_ser_solid_fill = true
            elsif @inside_drop_lines_sp_pr
              @inside_drop_lines_solid_fill = true
            elsif @inside_hi_low_lines_sp_pr
              @inside_hi_low_lines_solid_fill = true
            elsif @inside_ser_lines_sp_pr
              @inside_ser_lines_solid_fill = true
            elsif @inside_up_down_bar_sp_pr
              @inside_up_down_bar_solid_fill = true
            elsif @inside_ax_title_sp_pr
              @inside_ax_title_solid_fill = true
            elsif @inside_disp_units_lbl_sp_pr
              @inside_disp_units_lbl_solid_fill = true
            elsif @inside_plot_area_sp_pr
              @inside_plot_area_solid_fill = true
            elsif @inside_gridlines_sp_pr
              @inside_gridlines_solid_fill = true
            elsif @inside_ax_sp_pr
              @inside_ax_solid_fill = true
            elsif @inside_wall_sp_pr
              @inside_wall_solid_fill = true
            elsif @inside_legend_sp_pr
              @inside_legend_solid_fill = true
            elsif @inside_d_table_sp_pr
              @inside_d_table_solid_fill = true
            elsif @inside_title_sp_pr
              @inside_title_solid_fill = true
            elsif @inside_chart_space_sp_pr
              @inside_chart_space_solid_fill = true
            end
          when "noFill"
            if @inside_band_fmt_sp_pr && @current_band_fmt
              @current_band_fmt[:no_fill] = true
            elsif @inside_dpt_marker_sp_pr && @inside_dpt_marker_ln && @current_dpt
              @current_dpt[:marker_no_line] = true
            elsif @inside_dpt_marker_sp_pr && @current_dpt
              @current_dpt[:marker_no_fill] = true
            elsif @inside_ser && @inside_ser_ln && @current_ser
              @current_ser[:no_line] = true
            elsif @inside_marker_sp_pr && @inside_marker_ln && @current_ser
              @current_ser[:marker_no_line] = true
            elsif @inside_ser && @inside_ser_sp_pr && @current_ser
              @current_ser[:no_fill] = true
            elsif @inside_marker_sp_pr && @current_ser
              @current_ser[:marker_no_fill] = true
            elsif @inside_dpt && @inside_dpt_ln && @current_dpt
              @current_dpt[:no_line] = true
            elsif @inside_dpt && @inside_dpt_sp_pr && @current_dpt
              @current_dpt[:no_fill] = true
            elsif @inside_up_down_bar_sp_pr
              bar_key = @inside_up_bars ? :up_bars : :down_bars
              @up_down_bars[bar_key] ||= {}
              @up_down_bars[bar_key][:no_fill] = true
            elsif @inside_wall_sp_pr && @current_wall
              @current_wall[:no_fill] = true
            elsif @inside_ax_title_sp_pr && @inside_cat_ax
              @cat_axis_title_no_fill = true
            elsif @inside_ax_title_sp_pr && @inside_val_ax
              @val_axis_title_no_fill = true
            elsif @inside_title_sp_pr
              @title_no_fill = true
            elsif @inside_dlbl_sp_pr && @inside_dlbl_ln && @current_dlbl
              @current_dlbl[:no_line] = true
            elsif @inside_dlbl_sp_pr && @current_dlbl
              @current_dlbl[:no_fill] = true
            elsif @inside_dlbls_sp_pr
              @data_labels[:no_fill] = true
            elsif @inside_trendline_lbl_sp_pr && @current_trendline
              (@current_trendline[:label] ||= {})[:no_fill] = true
            elsif @inside_err_bars_sp_pr && @current_err_bars
              @current_err_bars[:no_fill] = true
            elsif @inside_disp_units_lbl_sp_pr
              du = @val_axis_disp_units
              du = {} unless du.is_a?(Hash)
              @val_axis_disp_units = du
              (du[:label] ||= {})[:no_fill] = true
            elsif @inside_legend_sp_pr
              @legend[:no_fill] = true
            elsif @inside_d_table_sp_pr && @data_table
              @data_table[:no_fill] = true
            elsif @inside_ax_sp_pr && @inside_cat_ax
              @cat_axis_no_fill = true
            elsif @inside_ax_sp_pr && @inside_val_ax
              @val_axis_no_fill = true
            elsif @inside_plot_area_sp_pr
              @plot_area_no_fill = true
            elsif @inside_chart_space_sp_pr
              @chart_no_fill = true
            end
          when "srgbClr"
            assign_chart_color(attributes["val"]) if attributes["val"]
          when "schemeClr"
            assign_chart_color({ scheme: attributes["val"] }) if attributes["val"]
          when "cat", "xVal"
            @inside_cat = true if @inside_ser
          when "val", "yVal"
            if @inside_err_bars && @current_err_bars && attributes["val"]
              @current_err_bars[:val] = attributes["val"].to_f
            elsif @inside_ser
              @inside_val = true
            end
          when "bubbleSize"
            @inside_bubble_size = true if @inside_ser
          when "f"
            @inside_f = true
            @text_buffer = +""
          when "numRef"
            @current_ser[:cat_ref_type] = :num if @inside_ser && @inside_cat && @current_ser
          when "numCache"
            @inside_num_cache = true
            @cache_values = []
          when "strCache"
            @inside_str_cache = true
            @cache_values = []
          when "pt"
            if (@inside_num_cache || @inside_str_cache) && attributes["idx"]
              @inside_cache_pt = true
              @current_cache_idx = attributes["idx"].to_i
            end
          when "v"
            if @inside_cache_pt
              @inside_cache_v = true
              @text_buffer = +""
            end
          when "title"
            @title_depth += 1
            if @inside_cat_ax || @inside_val_ax
              @inside_ax_title = true
            elsif @title_depth == 1
              @inside_title = true
            end
          when "t"
            @inside_t = true
            @text_buffer = +""
          when "rPr"
            if @inside_ax_title
              if @inside_cat_ax
                @cat_axis_title_font ||= {}
                @cat_axis_title_font[:bold] = true if attributes["b"] == "1"
                @cat_axis_title_font[:italic] = true if attributes["i"] == "1"
                @cat_axis_title_font[:size] = attributes["sz"].to_i if attributes["sz"]
              elsif @inside_val_ax
                @val_axis_title_font ||= {}
                @val_axis_title_font[:bold] = true if attributes["b"] == "1"
                @val_axis_title_font[:italic] = true if attributes["i"] == "1"
                @val_axis_title_font[:size] = attributes["sz"].to_i if attributes["sz"]
              end
              @inside_ax_title_rpr = true
            elsif @inside_title && @title_depth == 1
              @inside_title_rpr = true
              @title_font ||= {}
              @title_font[:bold] = true if attributes["b"] == "1"
              @title_font[:italic] = true if attributes["i"] == "1"
              @title_font[:size] = attributes["sz"].to_i if attributes["sz"]
            end
          when "latin"
            if @inside_axis_def_rpr && attributes["typeface"]
              if @inside_disp_units_lbl
                (@disp_units_lbl_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_cat_ax
                (@cat_axis_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_val_ax
                (@val_axis_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_legend_entry && @current_legend_entry
                (@current_legend_entry[:font] ||= {})[:name] = attributes["typeface"]
              elsif @inside_legend
                (@legend_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_d_table
                (@d_table_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_dlbl && @current_dlbl
                (@current_dlbl[:font] ||= {})[:name] = attributes["typeface"]
              elsif @inside_dlbls
                (@dlbls_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_trendline_lbl && @current_trendline
                (@trendline_lbl_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_chart_space_tx_pr
                (@chart_font ||= {})[:name] = attributes["typeface"]
              end
            elsif @inside_title_rpr && @title_font && attributes["typeface"]
              @title_font[:name] = attributes["typeface"]
            elsif @inside_ax_title_rpr && attributes["typeface"]
              if @inside_cat_ax
                (@cat_axis_title_font ||= {})[:name] = attributes["typeface"]
              elsif @inside_val_ax
                (@val_axis_title_font ||= {})[:name] = attributes["typeface"]
              end
            end
          when "legend"
            @inside_legend = true
          when "legendPos"
            @legend[:position] = attributes["val"] if @inside_legend && attributes["val"]
          when "manualLayout"
            @inside_legend_layout = true if @inside_legend
            @inside_plot_area_layout = true if @inside_plot_area && !@inside_legend && !@inside_title
            @inside_dlbl_layout = true if @inside_dlbl
            @inside_title_layout = true if (@inside_title || @inside_ax_title) && !@inside_legend && !@inside_dlbl
            @inside_trendline_lbl_layout = true if @inside_trendline_lbl
          when "layoutTarget"
            (@legend[:layout] ||= {})[:target] = attributes["val"] if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:target] = attributes["val"] if @inside_plot_area_layout && attributes["val"]
          when "xMode"
            (@legend[:layout] ||= {})[:x_mode] = attributes["val"] if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:x_mode] = attributes["val"] if @inside_plot_area_layout && attributes["val"]
          when "yMode"
            (@legend[:layout] ||= {})[:y_mode] = attributes["val"] if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:y_mode] = attributes["val"] if @inside_plot_area_layout && attributes["val"]
          when "wMode"
            (@legend[:layout] ||= {})[:w_mode] = attributes["val"] if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:w_mode] = attributes["val"] if @inside_plot_area_layout && attributes["val"]
          when "hMode"
            (@legend[:layout] ||= {})[:h_mode] = attributes["val"] if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:h_mode] = attributes["val"] if @inside_plot_area_layout && attributes["val"]
          when "x"
            (@legend[:layout] ||= {})[:x] = attributes["val"].to_f if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:x] = attributes["val"].to_f if @inside_plot_area_layout && attributes["val"]
            (@current_dlbl[:layout] ||= {})[:x] = attributes["val"].to_f if @inside_dlbl_layout && @current_dlbl && attributes["val"]
            assign_title_layout_value(:x, attributes["val"].to_f) if @inside_title_layout && attributes["val"]
            ((@current_trendline[:label] ||= {})[:layout] ||= {})[:x] = attributes["val"].to_f if @inside_trendline_lbl_layout && @current_trendline && attributes["val"]
          when "y"
            (@legend[:layout] ||= {})[:y] = attributes["val"].to_f if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:y] = attributes["val"].to_f if @inside_plot_area_layout && attributes["val"]
            (@current_dlbl[:layout] ||= {})[:y] = attributes["val"].to_f if @inside_dlbl_layout && @current_dlbl && attributes["val"]
            assign_title_layout_value(:y, attributes["val"].to_f) if @inside_title_layout && attributes["val"]
            ((@current_trendline[:label] ||= {})[:layout] ||= {})[:y] = attributes["val"].to_f if @inside_trendline_lbl_layout && @current_trendline && attributes["val"]
          when "w"
            (@legend[:layout] ||= {})[:w] = attributes["val"].to_f if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:w] = attributes["val"].to_f if @inside_plot_area_layout && attributes["val"]
            assign_title_layout_value(:w, attributes["val"].to_f) if @inside_title_layout && attributes["val"]
          when "h"
            (@legend[:layout] ||= {})[:h] = attributes["val"].to_f if @inside_legend_layout && attributes["val"]
            (@plot_area_layout ||= {})[:h] = attributes["val"].to_f if @inside_plot_area_layout && attributes["val"]
            assign_title_layout_value(:h, attributes["val"].to_f) if @inside_title_layout && attributes["val"]
          when "legendEntry"
            if @inside_legend
              @inside_legend_entry = true
              @current_legend_entry = {}
            end
          when "idx"
            if @inside_band_fmt && @current_band_fmt && attributes["val"]
              @current_band_fmt[:idx] = attributes["val"].to_i
            elsif @inside_dlbl && @current_dlbl && attributes["val"]
              @current_dlbl[:idx] = attributes["val"].to_i
            elsif @inside_dpt && @current_dpt && attributes["val"]
              @current_dpt[:idx] = attributes["val"].to_i
            elsif @inside_legend_entry && @current_legend_entry && attributes["val"]
              @current_legend_entry[:idx] = attributes["val"].to_i
            end
          when "delete"
            if @inside_dlbl && @current_dlbl && attributes["val"]
              @current_dlbl[:delete] = attributes["val"] == "1"
            elsif @inside_legend_entry && @current_legend_entry && attributes["val"]
              @current_legend_entry[:delete] = attributes["val"] == "1"
            elsif attributes["val"]
              if @inside_cat_ax
                @cat_axis_delete = attributes["val"] == "1"
              elsif @inside_val_ax
                @val_axis_delete = attributes["val"] == "1"
              end
            end
          when "overlay"
            if @inside_legend && attributes["val"]
              @legend[:overlay] = attributes["val"] == "1"
            elsif @inside_title && @title_depth == 1 && !@inside_ax_title && attributes["val"]
              @title_overlay = attributes["val"] == "1"
            end
          when "dLbls"
            if @inside_ser || @chart_type
              @inside_dlbls = true
              @dlbl_target = if @inside_ser && @current_ser
                               (@current_ser[:data_labels] ||= {})
                             else
                               @data_labels
                             end
            end
          when "dLbl"
            if @inside_dlbls
              @inside_dlbl = true
              @current_dlbl = {}
            end
          when "tx"
            @inside_dlbl_tx = true if @inside_dlbl
            @inside_trendline_lbl_tx = true if @inside_trendline_lbl
          when "showVal"
            if @inside_dlbl && @current_dlbl
              @current_dlbl[:show_val] = attributes["val"] == "1"
            elsif @inside_dlbls
              @dlbl_target[:show_val] = attributes["val"] == "1"
              @data_labels[:show_val] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "showCatName"
            if @inside_dlbl && @current_dlbl
              @current_dlbl[:show_cat_name] = attributes["val"] == "1"
            elsif @inside_dlbls
              @dlbl_target[:show_cat_name] = attributes["val"] == "1"
              @data_labels[:show_cat_name] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "showSerName"
            if @inside_dlbl && @current_dlbl
              @current_dlbl[:show_ser_name] = attributes["val"] == "1"
            elsif @inside_dlbls
              @dlbl_target[:show_ser_name] = attributes["val"] == "1"
              @data_labels[:show_ser_name] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "showPercent"
            if @inside_dlbl && @current_dlbl
              @current_dlbl[:show_percent] = attributes["val"] == "1"
            elsif @inside_dlbls
              @dlbl_target[:show_percent] = attributes["val"] == "1"
              @data_labels[:show_percent] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "showLegendKey"
            if @inside_dlbl && @current_dlbl
              @current_dlbl[:show_legend_key] = attributes["val"] == "1"
            elsif @inside_dlbls
              @dlbl_target[:show_legend_key] = attributes["val"] == "1"
              @data_labels[:show_legend_key] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "dLblPos"
            if @inside_dlbl && @current_dlbl && attributes["val"]
              @current_dlbl[:position] = attributes["val"]
            elsif @inside_dlbls && attributes["val"]
              @dlbl_target[:position] = attributes["val"]
              @data_labels[:position] = attributes["val"] if @dlbl_target != @data_labels
            end
          when "showBubbleSize"
            if @inside_dlbl && @current_dlbl
              @current_dlbl[:show_bubble_size] = attributes["val"] == "1"
            elsif @inside_dlbls
              @dlbl_target[:show_bubble_size] = attributes["val"] == "1"
              @data_labels[:show_bubble_size] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "separator"
            if @inside_dlbls
              @inside_separator = true
              @text_buffer = +""
            end
          when "showLeaderLines"
            if @inside_dlbls && attributes["val"]
              @dlbl_target[:show_leader_lines] = attributes["val"] == "1"
              @data_labels[:show_leader_lines] = attributes["val"] == "1" if @dlbl_target != @data_labels
            end
          when "leaderLines"
            @inside_leader_lines = true if @inside_dlbls
          when "catAx"
            @inside_cat_ax = true
          when "dateAx"
            @inside_cat_ax = true
            @cat_axis_type = :date
          when "valAx"
            @inside_val_ax = true
          when "scaling"
            @inside_scaling = true if @inside_cat_ax || @inside_val_ax
          when "logBase"
            if @inside_scaling && attributes["val"]
              if @inside_cat_ax
                @cat_axis_log_base = attributes["val"].to_f
              elsif @inside_val_ax
                @val_axis_log_base = attributes["val"].to_f
              end
            end
          when "max"
            if @inside_scaling && attributes["val"]
              if @inside_cat_ax
                @cat_axis_scaling_max = attributes["val"].to_f
              elsif @inside_val_ax
                @val_axis_scaling_max = attributes["val"].to_f
              end
            end
          when "min"
            if @inside_scaling && attributes["val"]
              if @inside_cat_ax
                @cat_axis_scaling_min = attributes["val"].to_f
              elsif @inside_val_ax
                @val_axis_scaling_min = attributes["val"].to_f
              end
            end
          when "orientation"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_orientation = attributes["val"]
              elsif @inside_val_ax
                @val_axis_orientation = attributes["val"]
              end
            end
          when "numFmt"
            if @inside_trendline_lbl && @current_trendline && attributes["formatCode"]
              nf = { format_code: attributes["formatCode"] }
              nf[:source_linked] = attributes["sourceLinked"] == "1" if attributes["sourceLinked"]
              (@current_trendline[:label] ||= {})[:num_fmt] = nf
            elsif @inside_disp_units_lbl && attributes["formatCode"]
              nf = { format_code: attributes["formatCode"] }
              nf[:source_linked] = attributes["sourceLinked"] == "1" if attributes["sourceLinked"]
              if @val_axis_disp_units.is_a?(String)
                @val_axis_disp_units = { built_in_unit: @val_axis_disp_units }
              elsif @val_axis_disp_units.nil?
                @val_axis_disp_units = {}
              end
              (@val_axis_disp_units[:label] ||= {})[:num_fmt] = nf
            elsif (@inside_cat_ax || @inside_val_ax) && attributes["formatCode"]
              nf = { format_code: attributes["formatCode"] }
              nf[:source_linked] = attributes["sourceLinked"] == "1" if attributes["sourceLinked"]
              if @inside_cat_ax
                @cat_axis_num_fmt = nf
              elsif @inside_val_ax
                @val_axis_num_fmt = nf
              end
            elsif @inside_dlbl && @current_dlbl && attributes["formatCode"]
              nf = { format_code: attributes["formatCode"] }
              nf[:source_linked] = attributes["sourceLinked"] == "1" if attributes["sourceLinked"]
              @current_dlbl[:num_fmt] = nf
            elsif @inside_dlbls && !@inside_dlbl && attributes["formatCode"]
              nf = { format_code: attributes["formatCode"] }
              nf[:source_linked] = attributes["sourceLinked"] == "1" if attributes["sourceLinked"]
              @dlbl_target[:num_fmt] = nf
              @data_labels[:num_fmt] = nf if @dlbl_target != @data_labels
            end
          when "majorTickMark"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_major_tick_mark = attributes["val"]
              elsif @inside_val_ax
                @val_axis_major_tick_mark = attributes["val"]
              end
            end
          when "minorTickMark"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_minor_tick_mark = attributes["val"]
              elsif @inside_val_ax
                @val_axis_minor_tick_mark = attributes["val"]
              end
            end
          when "crosses"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_crosses = attributes["val"]
              elsif @inside_val_ax
                @val_axis_crosses = attributes["val"]
              end
            end
          when "crossesAt"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_crosses_at = attributes["val"]&.to_f
              elsif @inside_val_ax
                @val_axis_crosses_at = attributes["val"]&.to_f
              end
            end
          when "tickLblSkip"
            @cat_axis_tick_lbl_skip = attributes["val"]&.to_i if attributes["val"] && @inside_cat_ax
          when "tickMarkSkip"
            @cat_axis_tick_mark_skip = attributes["val"]&.to_i if attributes["val"] && @inside_cat_ax
          when "auto"
            @cat_axis_auto = attributes["val"] == "1" if attributes["val"] && @inside_cat_ax
          when "lblAlgn"
            @cat_axis_lbl_algn = attributes["val"] if attributes["val"] && @inside_cat_ax
          when "lblOffset"
            @cat_axis_lbl_offset = attributes["val"]&.to_i if attributes["val"] && @inside_cat_ax
          when "noMultiLvlLbl"
            @cat_axis_no_multi_lvl_lbl = attributes["val"] == "1" if attributes["val"] && @inside_cat_ax
          when "axPos"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_pos = attributes["val"]
              elsif @inside_val_ax
                @val_axis_pos = attributes["val"]
              end
            end
          when "crossBetween"
            @val_axis_cross_between = attributes["val"] if @inside_val_ax && attributes["val"]
          when "majorUnit"
            if @inside_cat_ax && @cat_axis_type == :date && attributes["val"]
              @cat_axis_major_unit = attributes["val"].to_f
            elsif @inside_val_ax && attributes["val"]
              @val_axis_major_unit = attributes["val"].to_f
            end
          when "minorUnit"
            if @inside_cat_ax && @cat_axis_type == :date && attributes["val"]
              @cat_axis_minor_unit = attributes["val"].to_f
            elsif @inside_val_ax && attributes["val"]
              @val_axis_minor_unit = attributes["val"].to_f
            end
          when "baseTimeUnit"
            @cat_axis_base_time_unit = attributes["val"] if @inside_cat_ax && attributes["val"]
          when "majorTimeUnit"
            @cat_axis_major_time_unit = attributes["val"] if @inside_cat_ax && attributes["val"]
          when "minorTimeUnit"
            @cat_axis_minor_time_unit = attributes["val"] if @inside_cat_ax && attributes["val"]
          when "builtInUnit"
            @val_axis_disp_units = attributes["val"] if @inside_val_ax && @inside_disp_units && attributes["val"]
          when "custUnit"
            @val_axis_disp_units = { cust_unit: attributes["val"].to_f } if @inside_val_ax && @inside_disp_units && attributes["val"]
          when "dispUnits"
            @inside_disp_units = true if @inside_val_ax
          when "dispUnitsLbl"
            @inside_disp_units_lbl = true if @inside_disp_units
          when "txPr"
            if @inside_cat_ax || @inside_val_ax || @inside_legend || @inside_legend_entry || @inside_d_table || @inside_dlbls || @inside_trendline_lbl || @inside_disp_units_lbl
              @inside_axis_tx_pr = true
            elsif !@inside_chart
              @inside_chart_space_tx_pr = true
              @inside_axis_tx_pr = true
            end
          when "bodyPr"
            if attributes["rot"]
              if @inside_axis_tx_pr
                if @inside_cat_ax
                  @cat_axis_label_rotation = attributes["rot"].to_i
                elsif @inside_val_ax
                  @val_axis_label_rotation = attributes["rot"].to_i
                end
              elsif @inside_ax_title
                if @inside_cat_ax
                  @cat_axis_title_rotation = attributes["rot"].to_i
                elsif @inside_val_ax
                  @val_axis_title_rotation = attributes["rot"].to_i
                end
              elsif @inside_title && @title_depth == 1
                @title_rotation = attributes["rot"].to_i
              end
            end
          when "defRPr"
            if @inside_axis_tx_pr
              @inside_axis_def_rpr = true
              font = {}
              font[:size] = attributes["sz"].to_i / 100.0 if attributes["sz"]
              font[:bold] = true if attributes["b"] == "1"
              font[:italic] = true if attributes["i"] == "1"
              if @inside_disp_units_lbl
                @disp_units_lbl_font = (@disp_units_lbl_font || {}).merge(font)
              elsif @inside_cat_ax
                @cat_axis_font = (@cat_axis_font || {}).merge(font)
              elsif @inside_val_ax
                @val_axis_font = (@val_axis_font || {}).merge(font)
              elsif @inside_legend_entry && @current_legend_entry
                @current_legend_entry[:font] = (@current_legend_entry[:font] || {}).merge(font)
              elsif @inside_legend
                @legend_font = (@legend_font || {}).merge(font)
              elsif @inside_d_table
                @d_table_font = (@d_table_font || {}).merge(font)
              elsif @inside_dlbl && @current_dlbl
                @current_dlbl[:font] = (@current_dlbl[:font] || {}).merge(font)
              elsif @inside_dlbls
                @dlbls_font = (@dlbls_font || {}).merge(font)
              elsif @inside_trendline_lbl && @current_trendline
                @trendline_lbl_font = (@trendline_lbl_font || {}).merge(font)
              elsif @inside_chart_space_tx_pr
                @chart_font = (@chart_font || {}).merge(font)
              end
            end
          when "tickLblPos"
            if attributes["val"]
              if @inside_cat_ax
                @cat_axis_tick_lbl_pos = attributes["val"]
              elsif @inside_val_ax
                @val_axis_tick_lbl_pos = attributes["val"]
              end
            end
          when "majorGridlines"
            @inside_gridlines = true
            if @inside_cat_ax
              @cat_axis_major_gridlines = true
              @gridlines_target = :cat_axis_major_gridlines
            elsif @inside_val_ax
              @val_axis_major_gridlines = true
              @gridlines_target = :val_axis_major_gridlines
            end
          when "minorGridlines"
            @inside_gridlines = true
            if @inside_cat_ax
              @cat_axis_minor_gridlines = true
              @gridlines_target = :cat_axis_minor_gridlines
            elsif @inside_val_ax
              @val_axis_minor_gridlines = true
              @gridlines_target = :val_axis_minor_gridlines
            end
          when "plotVisOnly"
            @plot_vis_only = attributes["val"] == "1" if attributes["val"]
          when "dispBlanksAs"
            @disp_blanks_as = attributes["val"] if attributes["val"]
          when "style"
            @style = attributes["val"]&.to_i if attributes["val"]
          when "roundedCorners"
            @rounded_corners = attributes["val"] == "1" if attributes["val"]
          when "protection"
            unless @inside_chart
              @inside_protection = true
              @protection = {}
            end
          when "chartObject"
            @protection[:chart_object] = attributes["val"] == "1" if @inside_protection && attributes["val"]
          when "data"
            @protection[:data] = attributes["val"] == "1" if @inside_protection && attributes["val"]
          when "formatting"
            @protection[:formatting] = attributes["val"] == "1" if @inside_protection && attributes["val"]
          when "selection"
            @protection[:selection] = attributes["val"] == "1" if @inside_protection && attributes["val"]
          when "userInterface"
            @protection[:user_interface] = attributes["val"] == "1" if @inside_protection && attributes["val"]
          when "showDLblsOverMax"
            @show_d_lbls_over_max = attributes["val"] == "1" if attributes["val"]
          when "dTable"
            @inside_d_table = true
            @data_table = {}
          when "showHorzBorder"
            @data_table[:show_horz_border] = attributes["val"] == "1" if @inside_d_table && @data_table && attributes["val"]
          when "showVertBorder"
            @data_table[:show_vert_border] = attributes["val"] == "1" if @inside_d_table && @data_table && attributes["val"]
          when "showOutline"
            @data_table[:show_outline] = attributes["val"] == "1" if @inside_d_table && @data_table && attributes["val"]
          when "showKeys"
            @data_table[:show_keys] = attributes["val"] == "1" if @inside_d_table && @data_table && attributes["val"]
          when "printSettings"
            unless @inside_chart
              @inside_print_settings = true
              @print_settings = {}
            end
          when "headerFooter"
            if @inside_print_settings
              @inside_ps_header_footer = true
              @print_settings[:header_footer] = {}
            end
          when "oddHeader"
            @inside_ps_odd_header = true if @inside_ps_header_footer
            @text_buffer = +""
          when "oddFooter"
            @inside_ps_odd_footer = true if @inside_ps_header_footer
            @text_buffer = +""
          when "evenHeader"
            @inside_ps_even_header = true if @inside_ps_header_footer
            @text_buffer = +""
          when "evenFooter"
            @inside_ps_even_footer = true if @inside_ps_header_footer
            @text_buffer = +""
          when "firstHeader"
            @inside_ps_first_header = true if @inside_ps_header_footer
            @text_buffer = +""
          when "firstFooter"
            @inside_ps_first_footer = true if @inside_ps_header_footer
            @text_buffer = +""
          when "pageMargins"
            if @inside_print_settings
              pm = {}
              %w[b l r t header footer].each do |a|
                pm[a.to_sym] = attributes[a].to_f if attributes[a]
              end
              @print_settings[:page_margins] = pm unless pm.empty?
            end
          when "pageSetup"
            if @inside_print_settings
              psu = {}
              psu[:paper_size] = attributes["paperSize"].to_i if attributes["paperSize"]
              psu[:first_page_number] = attributes["firstPageNumber"].to_i if attributes["firstPageNumber"]
              psu[:orientation] = attributes["orientation"] if attributes["orientation"]
              psu[:horizontal_dpi] = attributes["horizontalDpi"].to_i if attributes["horizontalDpi"]
              psu[:vertical_dpi] = attributes["verticalDpi"].to_i if attributes["verticalDpi"]
              psu[:copies] = attributes["copies"].to_i if attributes["copies"]
              @print_settings[:page_setup] = psu unless psu.empty?
            end
          end
        end

        def characters(text)
          @text_buffer << text if @inside_t || @inside_f || @inside_separator || @inside_trendline_name || @inside_cache_v ||
                                  @inside_ps_odd_header || @inside_ps_odd_footer ||
                                  @inside_ps_even_header || @inside_ps_even_footer ||
                                  @inside_ps_first_header || @inside_ps_first_footer
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "t"
            if @inside_dlbl_tx && @current_dlbl
              @current_dlbl[:text] = (@current_dlbl[:text] || +"") << @text_buffer
            elsif @inside_trendline_lbl_tx && @current_trendline
              (@current_trendline[:label] ||= {})[:text] = (@current_trendline.dig(:label, :text) || +"") << @text_buffer
            elsif @inside_ax_title
              if @inside_cat_ax
                @cat_axis_title = @text_buffer.dup
              elsif @inside_val_ax
                @val_axis_title = @text_buffer.dup
              end
            elsif @inside_title && @title_depth == 1
              @title = @text_buffer.dup
            end
            @inside_t = false
          when "f"
            if @inside_err_bars && @current_err_bars
              if @inside_err_bars_plus
                @current_err_bars[:plus] = @text_buffer.dup
              elsif @inside_err_bars_minus
                @current_err_bars[:minus] = @text_buffer.dup
              end
            elsif @inside_ser
              if @inside_cat
                @current_ser[:cat_ref] = @text_buffer.dup
              elsif @inside_val
                @current_ser[:val_ref] = @text_buffer.dup
              elsif @inside_bubble_size
                @current_ser[:bubble_size_ref] = @text_buffer.dup
              else
                @current_ser[:name] = @text_buffer.dup
              end
            end
            @inside_f = false
          when "v"
            if @inside_cache_v
              @cache_values[@current_cache_idx] = @text_buffer.dup
              @inside_cache_v = false
            end
          when "pt"
            @inside_cache_pt = false
          when "numCache"
            if @inside_num_cache && @inside_ser && @current_ser
              if @inside_cat
                @current_ser[:cat_cache] = @cache_values.dup
              elsif @inside_val
                @current_ser[:val_cache] = @cache_values.dup
              elsif @inside_bubble_size
                @current_ser[:bubble_size_cache] = @cache_values.dup
              end
            end
            @inside_num_cache = false
            @cache_values = []
          when "strCache"
            if @inside_str_cache && @inside_ser && @current_ser
              if @inside_cat
                @current_ser[:cat_cache] = @cache_values.dup
              else
                @current_ser[:name_cache] = @cache_values.dup
              end
            end
            @inside_str_cache = false
            @cache_values = []
          when "cat", "xVal"
            @inside_cat = false
          when "val", "yVal"
            @inside_val = false
          when "bubbleSize"
            @inside_bubble_size = false
          when "dPt"
            if @inside_dpt && @current_dpt && @current_ser
              @current_ser[:data_points] ||= []
              @current_ser[:data_points] << @current_dpt
            end
            @current_dpt = nil
            @inside_dpt = false
            @inside_dpt_marker = false
            @inside_dpt_marker_sp_pr = false
            @inside_dpt_marker_solid_fill = false
            @inside_dpt_marker_ln = false
            @inside_dpt_sp_pr = false
            @inside_dpt_solid_fill = false
            @inside_dpt_ln = false
          when "trendline"
            if @inside_trendline && @current_trendline && @current_ser
              (@current_ser[:trendlines] ||= []) << @current_trendline
              @current_ser[:trendline] = @current_ser[:trendlines].first
            end
            @current_trendline = nil
            @inside_trendline = false
            @inside_trendline_name = false
            @inside_trendline_sp_pr = false
            @inside_trendline_ln = false
            @inside_trendline_solid_fill = false
            @inside_trendline_lbl = false
            @inside_trendline_lbl_layout = false
            @inside_trendline_lbl_tx = false
            @inside_trendline_lbl_sp_pr = false
            @inside_trendline_lbl_ln = false
            @inside_trendline_lbl_solid_fill = false
            @trendline_lbl_font = nil
          when "trendlineLbl"
            (@current_trendline[:label] ||= {})[:font] = @trendline_lbl_font if @trendline_lbl_font && @current_trendline
            @inside_trendline_lbl = false
            @inside_trendline_lbl_layout = false
            @inside_trendline_lbl_tx = false
            @inside_trendline_lbl_sp_pr = false
            @inside_trendline_lbl_ln = false
            @inside_trendline_lbl_solid_fill = false
            @trendline_lbl_font = nil
          when "errBars"
            if @inside_err_bars && @current_err_bars && @current_ser
              (@current_ser[:error_bars_list] ||= []) << @current_err_bars
              @current_ser[:error_bars] = @current_ser[:error_bars_list].first
            end
            @current_err_bars = nil
            @inside_err_bars = false
            @inside_err_bars_sp_pr = false
            @inside_err_bars_ln = false
            @inside_err_bars_solid_fill = false
            @inside_err_bars_plus = false
            @inside_err_bars_minus = false
          when "plus"
            @inside_err_bars_plus = false if @inside_err_bars
          when "minus"
            @inside_err_bars_minus = false if @inside_err_bars
          when "name"
            @current_trendline[:name] = @text_buffer.dup if @inside_trendline_name && @current_trendline
            @inside_trendline_name = false
          when "ser"
            @series << @current_ser if @current_ser
            @current_ser = nil
            @inside_ser = false
            @inside_ser_sp_pr = false
            @inside_ser_solid_fill = false
            @inside_ser_ln = false
            @inside_ser_marker = false
            @inside_marker_sp_pr = false
            @inside_marker_ln = false
            @inside_marker_solid_fill = false
          when "spPr"
            if @inside_band_fmt_sp_pr
              @inside_band_fmt_sp_pr = false
              @inside_band_fmt_ln = false
              @inside_band_fmt_solid_fill = false
            elsif @inside_dpt_marker_sp_pr
              @inside_dpt_marker_sp_pr = false
              @inside_dpt_marker_ln = false
              @inside_dpt_marker_solid_fill = false
            elsif @inside_dpt
              @inside_dpt_sp_pr = false
              @inside_dpt_solid_fill = false
            elsif @inside_marker_sp_pr
              @inside_marker_sp_pr = false
              @inside_marker_ln = false
              @inside_marker_solid_fill = false
            elsif @inside_leader_lines_sp_pr
              @inside_leader_lines_sp_pr = false
              @inside_leader_lines_ln = false
              @inside_leader_lines_solid_fill = false
            elsif @inside_dlbl_sp_pr
              @inside_dlbl_sp_pr = false
              @inside_dlbl_ln = false
              @inside_dlbl_solid_fill = false
            elsif @inside_dlbls_sp_pr
              @inside_dlbls_sp_pr = false
              @inside_dlbls_ln = false
              @inside_dlbls_solid_fill = false
            elsif @inside_trendline_lbl_sp_pr
              @inside_trendline_lbl_sp_pr = false
              @inside_trendline_lbl_ln = false
              @inside_trendline_lbl_solid_fill = false
            elsif @inside_trendline_sp_pr
              @inside_trendline_sp_pr = false
              @inside_trendline_ln = false
              @inside_trendline_solid_fill = false
            elsif @inside_err_bars_sp_pr
              @inside_err_bars_sp_pr = false
              @inside_err_bars_ln = false
              @inside_err_bars_solid_fill = false
            elsif @inside_drop_lines_sp_pr
              @inside_drop_lines_sp_pr = false
              @inside_drop_lines_ln = false
              @inside_drop_lines_solid_fill = false
            elsif @inside_hi_low_lines_sp_pr
              @inside_hi_low_lines_sp_pr = false
              @inside_hi_low_lines_ln = false
              @inside_hi_low_lines_solid_fill = false
            elsif @inside_ser_lines_sp_pr
              @inside_ser_lines_sp_pr = false
              @inside_ser_lines_ln = false
              @inside_ser_lines_solid_fill = false
            elsif @inside_up_down_bar_sp_pr
              @inside_up_down_bar_sp_pr = false
              @inside_up_down_bar_ln = false
              @inside_up_down_bar_solid_fill = false
            elsif @inside_ax_title_sp_pr
              @inside_ax_title_sp_pr = false
              @inside_ax_title_ln = false
              @inside_ax_title_solid_fill = false
            elsif @inside_ser
              @inside_ser_sp_pr = false
              @inside_ser_ln = false
            elsif @inside_gridlines
              @inside_gridlines_sp_pr = false
              @inside_gridlines_ln = false
              @inside_gridlines_solid_fill = false
            elsif @inside_disp_units_lbl_sp_pr
              @inside_disp_units_lbl_sp_pr = false
              @inside_disp_units_lbl_ln = false
              @inside_disp_units_lbl_solid_fill = false
            elsif @inside_cat_ax || @inside_val_ax
              @inside_ax_sp_pr = false
              @inside_ax_ln = false
              @inside_ax_solid_fill = false
            elsif @inside_wall
              @inside_wall_sp_pr = false
              @inside_wall_ln = false
              @inside_wall_solid_fill = false
            elsif @inside_legend
              @inside_legend_sp_pr = false
              @inside_legend_ln = false
              @inside_legend_solid_fill = false
            elsif @inside_d_table
              @inside_d_table_sp_pr = false
              @inside_d_table_ln = false
              @inside_d_table_solid_fill = false
            elsif @inside_title && @title_depth == 1
              @inside_title_sp_pr = false
              @inside_title_ln = false
              @inside_title_solid_fill = false
            elsif @inside_plot_area
              @inside_plot_area_sp_pr = false
              @inside_plot_area_solid_fill = false
              @inside_plot_area_ln = false
            elsif @inside_chart_space_sp_pr
              @inside_chart_space_sp_pr = false
              @inside_chart_space_ln = false
              @inside_chart_space_solid_fill = false
            end
          when "ln"
            @inside_band_fmt_ln = false if @inside_band_fmt_sp_pr
            @inside_dpt_marker_ln = false if @inside_dpt_marker_sp_pr
            @inside_dpt_ln = false if @inside_dpt
            @inside_marker_ln = false if @inside_marker_sp_pr
            @inside_leader_lines_ln = false if @inside_leader_lines_sp_pr
            @inside_dlbl_ln = false if @inside_dlbl_sp_pr
            @inside_dlbls_ln = false if @inside_dlbls_sp_pr
            @inside_trendline_lbl_ln = false if @inside_trendline_lbl_sp_pr
            @inside_trendline_ln = false if @inside_trendline_sp_pr
            @inside_err_bars_ln = false if @inside_err_bars_sp_pr
            @inside_ser_ln = false if @inside_ser
            @inside_drop_lines_ln = false if @inside_drop_lines_sp_pr
            @inside_hi_low_lines_ln = false if @inside_hi_low_lines_sp_pr
            @inside_ser_lines_ln = false if @inside_ser_lines_sp_pr
            @inside_up_down_bar_ln = false if @inside_up_down_bar_sp_pr
            @inside_ax_title_ln = false if @inside_ax_title_sp_pr
            @inside_disp_units_lbl_ln = false if @inside_disp_units_lbl_sp_pr
            @inside_gridlines_ln = false if @inside_gridlines_sp_pr
            @inside_ax_ln = false if @inside_ax_sp_pr
            @inside_wall_ln = false if @inside_wall_sp_pr
            @inside_legend_ln = false if @inside_legend_sp_pr
            @inside_d_table_ln = false if @inside_d_table_sp_pr
            @inside_title_ln = false if @inside_title_sp_pr
            @inside_plot_area_ln = false if @inside_plot_area_sp_pr
            @inside_chart_space_ln = false if @inside_chart_space_sp_pr
          when "marker"
            if @inside_dpt
              @inside_dpt_marker = false
            elsif @inside_ser
              @inside_ser_marker = false
            end
          when "rPr"
            @inside_title_rpr = false
            @inside_ax_title_rpr = false
          when "solidFill"
            if @inside_band_fmt_sp_pr
              @inside_band_fmt_solid_fill = false
            elsif @inside_dpt_marker_sp_pr
              @inside_dpt_marker_solid_fill = false
            elsif @inside_dpt
              @inside_dpt_solid_fill = false
            elsif @inside_marker_sp_pr
              @inside_marker_solid_fill = false
            elsif @inside_dlbl_sp_pr
              @inside_dlbl_solid_fill = false
            elsif @inside_trendline_lbl_sp_pr
              @inside_trendline_lbl_solid_fill = false
            elsif @inside_trendline_sp_pr
              @inside_trendline_solid_fill = false
            elsif @inside_err_bars_sp_pr
              @inside_err_bars_solid_fill = false
            elsif @inside_ser
              @inside_ser_solid_fill = false
            elsif @inside_drop_lines_sp_pr
              @inside_drop_lines_solid_fill = false
            elsif @inside_hi_low_lines_sp_pr
              @inside_hi_low_lines_solid_fill = false
            elsif @inside_ser_lines_sp_pr
              @inside_ser_lines_solid_fill = false
            elsif @inside_up_down_bar_sp_pr
              @inside_up_down_bar_solid_fill = false
            elsif @inside_ax_title_sp_pr
              @inside_ax_title_solid_fill = false
            elsif @inside_gridlines_sp_pr
              @inside_gridlines_solid_fill = false
            elsif @inside_disp_units_lbl_sp_pr
              @inside_disp_units_lbl_solid_fill = false
            elsif @inside_ax_sp_pr
              @inside_ax_solid_fill = false
            elsif @inside_wall_sp_pr
              @inside_wall_solid_fill = false
            elsif @inside_plot_area_sp_pr
              @inside_plot_area_solid_fill = false
            elsif @inside_chart_space_sp_pr
              @inside_chart_space_solid_fill = false
            end
          when "majorGridlines", "minorGridlines"
            @inside_gridlines = false
            @inside_gridlines_sp_pr = false
            @inside_gridlines_ln = false
            @inside_gridlines_solid_fill = false
            @gridlines_target = nil
          when "chart"
            @inside_chart = false
          when "plotArea"
            @inside_plot_area = false
            @inside_plot_area_layout = false
          when "title"
            @title_depth -= 1
            @inside_title = false if @title_depth.zero?
            @inside_title_layout = false
            @inside_ax_title = false
            @inside_ax_title_rpr = false
            @inside_ax_title_sp_pr = false
            @inside_ax_title_ln = false
            @inside_ax_title_solid_fill = false
            @inside_title_rpr = false
          when "legendEntry"
            if @inside_legend_entry && @current_legend_entry
              @legend[:entries] ||= []
              @legend[:entries] << @current_legend_entry
            end
            @current_legend_entry = nil
            @inside_legend_entry = false
            @inside_axis_tx_pr = false
            @inside_axis_def_rpr = false
          when "legend"
            @inside_legend = false
            @inside_legend_layout = false
            @inside_axis_tx_pr = false
            @inside_axis_def_rpr = false
          when "dLbls"
            if @dlbls_font
              dl_target = @dlbl_target || @data_labels
              dl_target[:font] = @dlbls_font
              @data_labels[:font] = @dlbls_font if dl_target != @data_labels
            end
            @inside_dlbls = false
            @dlbl_target = nil
            @inside_dlbls_sp_pr = false
            @inside_dlbls_ln = false
            @inside_dlbls_solid_fill = false
            @inside_leader_lines = false
            @inside_leader_lines_sp_pr = false
            @inside_leader_lines_ln = false
            @inside_leader_lines_solid_fill = false
            @dlbls_font = nil
            @inside_axis_tx_pr = false
            @inside_axis_def_rpr = false
          when "dLbl"
            if @inside_dlbl && @current_dlbl && @dlbl_target
              (@dlbl_target[:labels] ||= []) << @current_dlbl
              (@data_labels[:labels] ||= []) << @current_dlbl if @dlbl_target != @data_labels
            end
            @current_dlbl = nil
            @inside_dlbl = false
            @inside_dlbl_tx = false
            @inside_dlbl_sp_pr = false
            @inside_dlbl_solid_fill = false
            @inside_dlbl_ln = false
            @inside_dlbl_layout = false
          when "tx"
            @inside_dlbl_tx = false if @inside_dlbl
          when "separator"
            if @inside_separator
              if @inside_dlbl && @current_dlbl
                @current_dlbl[:separator] = @text_buffer.dup
              elsif @dlbl_target
                @dlbl_target[:separator] = @text_buffer.dup
                @data_labels[:separator] = @text_buffer.dup if @dlbl_target != @data_labels
              end
              @inside_separator = false
            end
          when "catAx", "dateAx"
            @inside_cat_ax = false
            @inside_axis_tx_pr = false
            @inside_axis_def_rpr = false
          when "valAx"
            @inside_val_ax = false
            @inside_axis_tx_pr = false
            @inside_axis_def_rpr = false
            @inside_disp_units = false
            @inside_disp_units_lbl = false
            @inside_disp_units_lbl_sp_pr = false
            @inside_disp_units_lbl_ln = false
            @inside_disp_units_lbl_solid_fill = false
            @disp_units_lbl_font = nil
          when "dispUnits"
            @inside_disp_units = false
            @inside_disp_units_lbl = false
            @inside_disp_units_lbl_sp_pr = false
            @inside_disp_units_lbl_ln = false
            @inside_disp_units_lbl_solid_fill = false
            @disp_units_lbl_font = nil
          when "dispUnitsLbl"
            if @disp_units_lbl_font
              du = @val_axis_disp_units
              du = {} unless du.is_a?(Hash)
              @val_axis_disp_units = du
              (du[:label] ||= {})[:font] = @disp_units_lbl_font
            end
            @inside_disp_units_lbl = false
            @inside_disp_units_lbl_sp_pr = false
            @inside_disp_units_lbl_ln = false
            @inside_disp_units_lbl_solid_fill = false
            @disp_units_lbl_font = nil
          when "txPr"
            @inside_axis_def_rpr = false if @inside_axis_tx_pr
            if @inside_chart_space_tx_pr
              @inside_axis_tx_pr = false
              @inside_chart_space_tx_pr = false
            end
          when "scaling"
            @inside_scaling = false
          when "view3D"
            @inside_view_3d = false
          when "floor", "sideWall", "backWall"
            @inside_wall = false
            @current_wall = nil
            @inside_wall_sp_pr = false
            @inside_wall_ln = false
            @inside_wall_solid_fill = false
          when "dTable"
            @data_table[:font] = @d_table_font if @d_table_font && @data_table
            @inside_d_table = false
            @inside_d_table_sp_pr = false
            @inside_d_table_ln = false
            @inside_d_table_solid_fill = false
            @inside_axis_tx_pr = false
            @inside_axis_def_rpr = false
          when "protection"
            @inside_protection = false
          when "printSettings"
            @inside_print_settings = false
          when "headerFooter"
            @inside_ps_header_footer = false if @inside_print_settings
          when "oddHeader"
            if @inside_ps_odd_header
              @print_settings[:header_footer][:odd_header] = @text_buffer.dup unless @text_buffer.empty?
              @inside_ps_odd_header = false
            end
          when "oddFooter"
            if @inside_ps_odd_footer
              @print_settings[:header_footer][:odd_footer] = @text_buffer.dup unless @text_buffer.empty?
              @inside_ps_odd_footer = false
            end
          when "evenHeader"
            if @inside_ps_even_header
              @print_settings[:header_footer][:even_header] = @text_buffer.dup unless @text_buffer.empty?
              @inside_ps_even_header = false
            end
          when "evenFooter"
            if @inside_ps_even_footer
              @print_settings[:header_footer][:even_footer] = @text_buffer.dup unless @text_buffer.empty?
              @inside_ps_even_footer = false
            end
          when "firstHeader"
            if @inside_ps_first_header
              @print_settings[:header_footer][:first_header] = @text_buffer.dup unless @text_buffer.empty?
              @inside_ps_first_header = false
            end
          when "firstFooter"
            if @inside_ps_first_footer
              @print_settings[:header_footer][:first_footer] = @text_buffer.dup unless @text_buffer.empty?
              @inside_ps_first_footer = false
            end
          when "upDownBars"
            @inside_up_down_bars = false
            @inside_up_bars = false
            @inside_down_bars = false
            @inside_up_down_bar_sp_pr = false
            @inside_up_down_bar_ln = false
            @inside_up_down_bar_solid_fill = false
          when "upBars"
            @inside_up_bars = false
            @inside_up_down_bar_sp_pr = false
            @inside_up_down_bar_ln = false
            @inside_up_down_bar_solid_fill = false
          when "downBars"
            @inside_down_bars = false
            @inside_up_down_bar_sp_pr = false
            @inside_up_down_bar_ln = false
            @inside_up_down_bar_solid_fill = false
          when "dropLines"
            @inside_drop_lines = false
            @inside_drop_lines_sp_pr = false
            @inside_drop_lines_ln = false
            @inside_drop_lines_solid_fill = false
          when "hiLowLines"
            @inside_hi_low_lines = false
            @inside_hi_low_lines_sp_pr = false
            @inside_hi_low_lines_ln = false
            @inside_hi_low_lines_solid_fill = false
          when "serLines"
            @inside_ser_lines = false
            @inside_ser_lines_sp_pr = false
            @inside_ser_lines_ln = false
            @inside_ser_lines_solid_fill = false
          when "custSplit"
            @inside_cust_split = false
          when "bandFmt"
            if @inside_band_fmt && @current_band_fmt
              @band_fmts << @current_band_fmt
              @current_band_fmt = nil
              @inside_band_fmt = false
              @inside_band_fmt_sp_pr = false
              @inside_band_fmt_ln = false
              @inside_band_fmt_solid_fill = false
            end
          when "bandFmts"
            @inside_band_fmts = false
          when "leaderLines"
            @inside_leader_lines = false
            @inside_leader_lines_sp_pr = false
            @inside_leader_lines_ln = false
            @inside_leader_lines_solid_fill = false
          end
        end

        private

        def assign_title_layout_value(key, value)
          if @inside_ax_title
            if @inside_cat_ax
              (@cat_axis_title_layout ||= {})[key] = value
            elsif @inside_val_ax
              (@val_axis_title_layout ||= {})[key] = value
            end
          elsif @inside_title && @title_depth == 1
            (@title_layout ||= {})[key] = value
          end
        end

        def assign_chart_color(color_value)
          if @inside_axis_def_rpr
            if @inside_disp_units_lbl
              (@disp_units_lbl_font ||= {})[:color] = color_value
            elsif @inside_cat_ax
              (@cat_axis_font ||= {})[:color] = color_value
            elsif @inside_val_ax
              (@val_axis_font ||= {})[:color] = color_value
            elsif @inside_legend_entry && @current_legend_entry
              (@current_legend_entry[:font] ||= {})[:color] = color_value
            elsif @inside_legend
              (@legend_font ||= {})[:color] = color_value
            elsif @inside_d_table
              (@d_table_font ||= {})[:color] = color_value
            elsif @inside_dlbl && @current_dlbl
              (@current_dlbl[:font] ||= {})[:color] = color_value
            elsif @inside_dlbls
              (@dlbls_font ||= {})[:color] = color_value
            elsif @inside_trendline_lbl && @current_trendline
              (@trendline_lbl_font ||= {})[:color] = color_value
            elsif @inside_chart_space_tx_pr
              (@chart_font ||= {})[:color] = color_value
            end
          elsif @inside_title_rpr && @title_font
            @title_font[:color] = color_value
          elsif @inside_ax_title_rpr
            if @inside_cat_ax
              (@cat_axis_title_font ||= {})[:color] = color_value
            elsif @inside_val_ax
              (@val_axis_title_font ||= {})[:color] = color_value
            end
          elsif @inside_ax_title_sp_pr && @inside_ax_title_ln && @inside_ax_title_solid_fill
            if @inside_cat_ax
              @cat_axis_title_line_color = color_value
            elsif @inside_val_ax
              @val_axis_title_line_color = color_value
            end
          elsif @inside_ax_title_sp_pr && @inside_ax_title_solid_fill
            if @inside_cat_ax
              @cat_axis_title_fill = color_value
            elsif @inside_val_ax
              @val_axis_title_fill = color_value
            end
          elsif @inside_band_fmt_sp_pr && @inside_band_fmt_ln && @inside_band_fmt_solid_fill && @current_band_fmt
            @current_band_fmt[:line_color] = color_value
          elsif @inside_band_fmt_sp_pr && @inside_band_fmt_solid_fill && @current_band_fmt
            @current_band_fmt[:fill_color] = color_value
          elsif @inside_dpt_marker_sp_pr && @inside_dpt_marker_ln && @inside_dpt_marker_solid_fill && @current_dpt
            @current_dpt[:marker_line_color] = color_value
          elsif @inside_dpt_marker_sp_pr && @inside_dpt_marker_solid_fill && @current_dpt
            @current_dpt[:marker_fill] = color_value
          elsif @inside_dpt && @inside_dpt_sp_pr && @inside_dpt_ln && @current_dpt
            @current_dpt[:line_color] = color_value
          elsif @inside_dpt && @inside_dpt_sp_pr && @inside_dpt_solid_fill && @current_dpt
            @current_dpt[:fill_color] = color_value
          elsif @inside_marker_sp_pr && @inside_marker_ln && @inside_marker_solid_fill && @current_ser
            @current_ser[:marker_line_color] = color_value
          elsif @inside_marker_sp_pr && @inside_marker_solid_fill && @current_ser
            @current_ser[:marker_fill] = color_value
          elsif @inside_leader_lines_sp_pr && @inside_leader_lines_ln && @inside_leader_lines_solid_fill
            ll = (@dlbl_target[:leader_lines] ||= {})
            ll[:line_color] = color_value
            if @dlbl_target != @data_labels
              ll2 = (@data_labels[:leader_lines] ||= {})
              ll2[:line_color] = color_value
            end
          elsif @inside_dlbl_sp_pr && @inside_dlbl_ln && @inside_dlbl_solid_fill && @current_dlbl
            @current_dlbl[:line_color] = color_value
          elsif @inside_dlbl_sp_pr && @inside_dlbl_solid_fill && @current_dlbl
            @current_dlbl[:fill_color] = color_value
          elsif @inside_dlbls_sp_pr && @inside_dlbls_ln && @inside_dlbls_solid_fill
            dl_target = @dlbl_target || @data_labels
            dl_target[:line_color] = color_value
            @data_labels[:line_color] = color_value if dl_target != @data_labels
          elsif @inside_dlbls_sp_pr && @inside_dlbls_solid_fill
            dl_target = @dlbl_target || @data_labels
            dl_target[:fill_color] = color_value
            @data_labels[:fill_color] = color_value if dl_target != @data_labels
          elsif @inside_trendline_lbl_sp_pr && @inside_trendline_lbl_ln && @inside_trendline_lbl_solid_fill && @current_trendline
            (@current_trendline[:label] ||= {})[:line_color] = color_value
          elsif @inside_trendline_lbl_sp_pr && @inside_trendline_lbl_solid_fill && @current_trendline
            (@current_trendline[:label] ||= {})[:fill_color] = color_value
          elsif @inside_trendline_sp_pr && @inside_trendline_ln && @inside_trendline_solid_fill && @current_trendline
            @current_trendline[:line_color] = color_value
          elsif @inside_err_bars_sp_pr && @inside_err_bars_ln && @inside_err_bars_solid_fill && @current_err_bars
            @current_err_bars[:line_color] = color_value
          elsif @inside_err_bars_sp_pr && @inside_err_bars_solid_fill && @current_err_bars
            @current_err_bars[:fill_color] = color_value
          elsif @inside_ser && @inside_ser_sp_pr && @inside_ser_ln && @inside_ser_solid_fill && @current_ser
            @current_ser[:line_color] = color_value
          elsif @inside_ser && @inside_ser_sp_pr && @inside_ser_solid_fill && @current_ser
            @current_ser[:fill_color] = color_value
          elsif @inside_drop_lines_sp_pr && @inside_drop_lines_ln && @inside_drop_lines_solid_fill
            @drop_lines = {} if @drop_lines == true
            @drop_lines[:line_color] = color_value
          elsif @inside_hi_low_lines_sp_pr && @inside_hi_low_lines_ln && @inside_hi_low_lines_solid_fill
            @hi_low_lines = {} if @hi_low_lines == true
            @hi_low_lines[:line_color] = color_value
          elsif @inside_ser_lines_sp_pr && @inside_ser_lines_ln && @inside_ser_lines_solid_fill
            @ser_lines = {} if @ser_lines == true
            @ser_lines[:line_color] = color_value
          elsif @inside_up_down_bar_sp_pr && @inside_up_down_bar_ln && @inside_up_down_bar_solid_fill
            bar_key = @inside_up_bars ? :up_bars : :down_bars
            @up_down_bars[bar_key] ||= {}
            @up_down_bars[bar_key][:line_color] = color_value
          elsif @inside_up_down_bar_sp_pr && @inside_up_down_bar_solid_fill
            bar_key = @inside_up_bars ? :up_bars : :down_bars
            @up_down_bars[bar_key] ||= {}
            @up_down_bars[bar_key][:fill_color] = color_value
          elsif @inside_plot_area_sp_pr && @inside_plot_area_ln && @inside_plot_area_solid_fill
            @plot_area_line_color = color_value
          elsif @inside_plot_area_sp_pr && @inside_plot_area_solid_fill
            @plot_area_fill = color_value
          elsif @inside_gridlines_sp_pr && @inside_gridlines_ln && @inside_gridlines_solid_fill && @gridlines_target
            gl = instance_variable_get(:"@#{@gridlines_target}")
            gl = {} if gl == true
            gl[:line_color] = color_value
            instance_variable_set(:"@#{@gridlines_target}", gl)
          elsif @inside_ax_sp_pr && @inside_ax_ln && @inside_ax_solid_fill
            if @inside_cat_ax
              @cat_axis_line_color = color_value
            elsif @inside_val_ax
              @val_axis_line_color = color_value
            end
          elsif @inside_ax_sp_pr && @inside_ax_solid_fill
            if @inside_cat_ax
              @cat_axis_fill = color_value
            elsif @inside_val_ax
              @val_axis_fill = color_value
            end
          elsif @inside_disp_units_lbl_sp_pr && @inside_disp_units_lbl_ln && @inside_disp_units_lbl_solid_fill
            du = @val_axis_disp_units
            du = {} unless du.is_a?(Hash)
            @val_axis_disp_units = du
            (du[:label] ||= {})[:line_color] = color_value
          elsif @inside_disp_units_lbl_sp_pr && @inside_disp_units_lbl_solid_fill
            du = @val_axis_disp_units
            du = {} unless du.is_a?(Hash)
            @val_axis_disp_units = du
            (du[:label] ||= {})[:fill_color] = color_value
          elsif @inside_wall_sp_pr && @inside_wall_ln && @inside_wall_solid_fill && @current_wall
            @current_wall[:line_color] = color_value
          elsif @inside_wall_sp_pr && @inside_wall_solid_fill && @current_wall
            @current_wall[:fill_color] = color_value
          elsif @inside_legend_sp_pr && @inside_legend_ln && @inside_legend_solid_fill
            @legend[:line_color] = color_value
          elsif @inside_legend_sp_pr && @inside_legend_solid_fill
            @legend[:fill_color] = color_value
          elsif @inside_d_table_sp_pr && @inside_d_table_ln && @inside_d_table_solid_fill && @data_table
            @data_table[:line_color] = color_value
          elsif @inside_d_table_sp_pr && @inside_d_table_solid_fill && @data_table
            @data_table[:fill_color] = color_value
          elsif @inside_title_sp_pr && @inside_title_ln && @inside_title_solid_fill
            @title_line_color = color_value
          elsif @inside_title_sp_pr && @inside_title_solid_fill
            @title_fill_color = color_value
          elsif @inside_chart_space_sp_pr && @inside_chart_space_ln && @inside_chart_space_solid_fill
            @chart_line_color = color_value
          elsif @inside_chart_space_sp_pr && @inside_chart_space_solid_fill
            @chart_fill = color_value
          end
        end

        def element_name(local_name, qname)
          if local_name.nil? || local_name.empty?
            qname.to_s.split(":").last
          else
            local_name
          end
        end
      end

      # SAX2 listener for parsing comments XML.
      class CommentsListener
        include REXML::SAX2Listener

        attr_reader :comments

        def initialize
          @comments = []
          @authors = []
          @inside_authors = false
          @inside_author = false
          @inside_comment = false
          @inside_text = false
          @inside_r = false
          @inside_rpr = false
          @inside_t = false
          @current_comment = nil
          @text_buffer = +""
          @runs = []
          @current_font = {}
          @has_runs = false
        end

        def start_element(_uri, local_name, qname, attributes)
          name = element_name(local_name, qname)
          case name
          when "authors"
            @inside_authors = true
          when "author"
            @inside_author = true
            @text_buffer = +""
          when "comment"
            @inside_comment = true
            @current_comment = { ref: attributes["ref"], author_id: attributes["authorId"]&.to_i }
            @current_comment[:guid] = attributes["guid"] if attributes["guid"]
            sid = attributes["shapeId"]
            @current_comment[:shape_id] = sid.to_i if sid
          when "text"
            if @inside_comment
              @inside_text = true
              @text_buffer = +""
              @runs = []
              @has_runs = false
            end
          when "r"
            if @inside_text
              @inside_r = true
              @has_runs = true
              @current_font = {}
            end
          when "rPr"
            @inside_rpr = true if @inside_r
          when "b"
            @current_font[:bold] = true if @inside_rpr
          when "i"
            @current_font[:italic] = true if @inside_rpr
          when "strike"
            @current_font[:strike] = true if @inside_rpr
          when "u"
            if @inside_rpr
              val = attributes["val"]
              @current_font[:underline] = val || true
            end
          when "vertAlign"
            @current_font[:vert_align] = attributes["val"] if @inside_rpr && attributes["val"]
          when "sz"
            @current_font[:sz] = attributes["val"]&.to_f if @inside_rpr
          when "color"
            if @inside_rpr
              if attributes["rgb"]
                @current_font[:color] = attributes["rgb"]
              elsif attributes["theme"]
                @current_font[:theme] = attributes["theme"].to_i
                @current_font[:tint] = attributes["tint"].to_f if attributes["tint"]
              elsif attributes["indexed"]
                @current_font[:indexed] = attributes["indexed"].to_i
              end
            end
          when "rFont"
            @current_font[:name] = attributes["val"] if @inside_rpr
          when "family"
            @current_font[:family] = attributes["val"]&.to_i if @inside_rpr
          when "scheme"
            @current_font[:scheme] = attributes["val"] if @inside_rpr
          when "t"
            if @inside_text
              @inside_t = true
              @text_buffer = +"" if @inside_r
            end
          end
        end

        def characters(text)
          @text_buffer << text if @inside_author || @inside_t
        end

        def end_element(_uri, local_name, qname)
          name = element_name(local_name, qname)
          case name
          when "authors"
            @inside_authors = false
          when "author"
            @authors << @text_buffer.dup if @inside_authors
            @inside_author = false
          when "comment"
            if @current_comment
              aid = @current_comment[:author_id]
              @current_comment[:author] = @authors[aid] if aid && aid < @authors.size
              @current_comment.delete(:author_id)
              @comments << @current_comment
            end
            @inside_comment = false
            @current_comment = nil
          when "text"
            if @current_comment && @inside_text
              if @has_runs && @runs.any? { |r| r[:font] }
                @current_comment[:text] = Xlsxrb::Elements::RichText.new(runs: @runs)
              else
                plain = @has_runs ? @runs.map { |r| r[:text] }.join : @text_buffer.dup
                @current_comment[:text] = plain
              end
            end
            @inside_text = false
          when "t"
            @inside_t = false
          when "rPr"
            @inside_rpr = false
          when "r"
            if @inside_r
              run = { text: @text_buffer.dup }
              run[:font] = @current_font.dup unless @current_font.empty?
              @runs << run
              @inside_r = false
            end
          end
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
