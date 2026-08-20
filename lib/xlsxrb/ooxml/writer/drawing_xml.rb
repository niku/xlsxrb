# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    class Writer
      # Mixin containing DrawingML, Chart, Shape, Comment, and VML XML generation logic.
      module DrawingXml
        # : (untyped drawing_parts) -> untyped
        def generate_drawing_xml(drawing_parts)
          parts = [
            XML_HEADER,
            %(<xdr:wsDr xmlns:xdr="#{XDR_NS}" xmlns:a="#{A_NS}" xmlns:r="#{DOC_REL_NS}">)
          ]

          drawing_parts.each do |dp|
            case dp[:kind]
            when :pic
              img = dp[:img]
              rid = "rId#{dp[:rid_index]}"
              ea = img[:edit_as] || "oneCell"
              img_pub_attr = img[:published] ? ' fPublished="1"' : ""
              parts << %(<xdr:twoCellAnchor editAs="#{xml_escape(ea)}"#{img_pub_attr}>)
              parts << anchor_xml("from", img[:from_col], img[:from_row], col_off: img[:from_col_off] || 0, row_off: img[:from_row_off] || 0)
              parts << anchor_xml("to", img[:to_col], img[:to_row], col_off: img[:to_col_off] || 0, row_off: img[:to_row_off] || 0)
              macro_attr = img[:macro] ? %( macro="#{xml_escape(img[:macro])}") : ""
              parts << "<xdr:pic#{macro_attr}>"
              descr_attr = img[:description] ? %( descr="#{xml_escape(img[:description])}") : ""
              title_attr = img[:title] ? %( title="#{xml_escape(img[:title])}") : ""
              hidden_attr = img[:hidden] ? ' hidden="1"' : ""
              pic_lock_attrs = +""
              pic_lock_attrs << ' noChangeAspect="1"' if img[:no_change_aspect]
              pic_lock_attrs << ' noCrop="1"' if img[:no_crop]
              pic_locks = pic_lock_attrs.empty? ? "<a:picLocks/>" : "<a:picLocks#{pic_lock_attrs}/>"
              parts << %(<xdr:nvPicPr><xdr:cNvPr id="#{dp[:rid_index] + 1}" name="#{xml_escape(img[:name])}"#{descr_attr}#{title_attr}#{hidden_attr}/><xdr:cNvPicPr>#{pic_locks}</xdr:cNvPicPr></xdr:nvPicPr>)
              src_rect_xml = if img[:src_rect]
                               sr = img[:src_rect]
                               sr_attrs = +""
                               sr_attrs << %( t="#{sr[:top]}") if sr[:top]
                               sr_attrs << %( b="#{sr[:bottom]}") if sr[:bottom]
                               sr_attrs << %( l="#{sr[:left]}") if sr[:left]
                               sr_attrs << %( r="#{sr[:right]}") if sr[:right]
                               "<a:srcRect#{sr_attrs}/>"
                             else
                               ""
                             end
              blip_xml = if img[:alpha_mod_fix]
                           %(<a:blip r:embed="#{rid}"><a:alphaModFix amt="#{img[:alpha_mod_fix]}"/></a:blip>)
                         else
                           %(<a:blip r:embed="#{rid}"/>)
                         end
              parts << %(<xdr:blipFill>#{blip_xml}#{src_rect_xml}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>)
              img_line_xml = if img[:line_color]
                               ln_w_attr = img[:line_width] ? %( w="#{img[:line_width].to_i}") : ""
                               %(<a:ln#{ln_w_attr}><a:solidFill>#{color_xml(img[:line_color])}</a:solidFill></a:ln>)
                             else
                               ""
                             end
              img_xfrm_xml = img[:rotation] ? %(<a:xfrm rot="#{img[:rotation].to_i}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>) : ""
              parts << %(<xdr:spPr>#{img_xfrm_xml}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>#{img_line_xml}</xdr:spPr>)
              parts << "</xdr:pic>"
              parts << client_data_xml(img)
              parts << "</xdr:twoCellAnchor>"
            when :chart
              chart = dp[:chart]
              rid = "rId#{dp[:rid_index]}"
              chart_ea_attr = chart[:edit_as] ? %( editAs="#{xml_escape(chart[:edit_as])}") : ""
              chart_pub_attr = chart[:published] ? ' fPublished="1"' : ""
              parts << "<xdr:twoCellAnchor#{chart_ea_attr}#{chart_pub_attr}>"
              parts << anchor_xml("from", chart[:from_col], chart[:from_row], col_off: chart[:from_col_off] || 0, row_off: chart[:from_row_off] || 0)
              parts << anchor_xml("to", chart[:to_col], chart[:to_row], col_off: chart[:to_col_off] || 0, row_off: chart[:to_row_off] || 0)
              gf_macro = chart[:frame_macro] ? xml_escape(chart[:frame_macro]) : ""
              parts << %(<xdr:graphicFrame macro="#{gf_macro}">)
              chart_frame_name = chart[:name] || chart[:title] || "Chart"
              chart_descr_attr = chart[:description] ? %( descr="#{xml_escape(chart[:description])}") : ""
              chart_title_attr = chart[:frame_title] ? %( title="#{xml_escape(chart[:frame_title])}") : ""
              chart_hidden_attr = chart[:frame_hidden] ? ' hidden="1"' : ""
              cnv_gf_pr = if chart[:frame_no_grp]
                            '<xdr:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></xdr:cNvGraphicFramePr>'
                          else
                            "<xdr:cNvGraphicFramePr/>"
                          end
              parts << %(<xdr:nvGraphicFramePr><xdr:cNvPr id="#{dp[:rid_index] + 1}" name="#{xml_escape(chart_frame_name)}"#{chart_descr_attr}#{chart_title_attr}#{chart_hidden_attr}/>#{cnv_gf_pr}</xdr:nvGraphicFramePr>)
              parts << '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="3000000"/></xdr:xfrm>'
              parts << %(<a:graphic><a:graphicData uri="#{C_NS}"><c:chart xmlns:c="#{C_NS}" xmlns:r="#{DOC_REL_NS}" r:id="#{rid}"/></a:graphicData></a:graphic>)
              parts << "</xdr:graphicFrame>"
              parts << client_data_xml(chart)
              parts << "</xdr:twoCellAnchor>"
            when :sp
              shape = dp[:shape]
              shape_ea_attr = shape[:edit_as] ? %( editAs="#{xml_escape(shape[:edit_as])}") : ""
              shape_pub_attr = shape[:published] ? ' fPublished="1"' : ""
              parts << "<xdr:twoCellAnchor#{shape_ea_attr}#{shape_pub_attr}>"
              parts << anchor_xml("from", shape[:from_col], shape[:from_row], col_off: shape[:from_col_off] || 0, row_off: shape[:from_row_off] || 0)
              parts << anchor_xml("to", shape[:to_col], shape[:to_row], col_off: shape[:to_col_off] || 0, row_off: shape[:to_row_off] || 0)
              sp_macro_attr = shape[:macro] ? %( macro="#{xml_escape(shape[:macro])}") : ""
              sp_textlink_attr = shape[:textlink] ? %( textlink="#{xml_escape(shape[:textlink])}") : ""
              parts << "<xdr:sp#{sp_macro_attr}#{sp_textlink_attr}>"
              shape_descr_attr = shape[:description] ? %( descr="#{xml_escape(shape[:description])}") : ""
              shape_title_attr = shape[:title] ? %( title="#{xml_escape(shape[:title])}") : ""
              shape_hidden_attr = shape[:hidden] ? ' hidden="1"' : ""
              sp_lock_attrs = +""
              sp_lock_attrs << ' noGrp="1"' if shape[:no_grp]
              sp_lock_attrs << ' noRot="1"' if shape[:no_rot]
              sp_lock_attrs << ' fLocksText="1"' if shape[:f_locks_text]
              cnv_sp_pr = if sp_lock_attrs.empty?
                            "<xdr:cNvSpPr/>"
                          else
                            "<xdr:cNvSpPr><a:spLocks#{sp_lock_attrs}/></xdr:cNvSpPr>"
                          end
              parts << %(<xdr:nvSpPr><xdr:cNvPr id="#{dp[:id]}" name="#{xml_escape(shape[:name])}"#{shape_descr_attr}#{shape_title_attr}#{shape_hidden_attr}/>#{cnv_sp_pr}</xdr:nvSpPr>)
              shape_fill_xml = if shape[:no_fill]
                                 "<a:noFill/>"
                               elsif shape[:fill_color]
                                 "<a:solidFill>#{srgb_clr_xml(shape[:fill_color], alpha: shape[:fill_alpha], transforms: shape[:fill_color_transforms])}</a:solidFill>"
                               elsif shape[:gradient_fill]
                                 gf = shape[:gradient_fill]
                                 gf_attrs = +""
                                 gf_attrs << %( rotWithShape="#{gf[:rot_with_shape] ? 1 : 0}") unless gf[:rot_with_shape].nil?
                                 gf_attrs << %( flip="#{xml_escape(gf[:flip])}") if gf[:flip]
                                 gs_xml = (gf[:stops] || []).map { |gs| %(<a:gs pos="#{gs[:pos]}">#{color_xml(gs[:color])}</a:gs>) }.join
                                 type_xml = if gf[:path]
                                              %(<a:path path="#{xml_escape(gf[:path])}"/>)
                                            elsif gf[:angle]
                                              scaled_attr = gf[:scaled] ? %( scaled="1") : %( scaled="0")
                                              %(<a:lin ang="#{gf[:angle]}"#{scaled_attr}/>)
                                            else
                                              ""
                                            end
                                 tile_xml = gf[:tile_rect] ? %(<a:tileRect#{gf[:tile_rect].map { |k, v| %( #{k}="#{v}") }.join}/>) : ""
                                 "<a:gradFill#{gf_attrs}><a:gsLst>#{gs_xml}</a:gsLst>#{type_xml}#{tile_xml}</a:gradFill>"
                               elsif shape[:pattern_fill]
                                 pf = shape[:pattern_fill]
                                 pf_children = +""
                                 pf_children << %(<a:fgClr>#{color_xml(pf[:fg_color])}</a:fgClr>) if pf[:fg_color]
                                 pf_children << %(<a:bgClr>#{color_xml(pf[:bg_color])}</a:bgClr>) if pf[:bg_color]
                                 %(<a:pattFill prst="#{xml_escape(pf[:preset])}">#{pf_children}</a:pattFill>)
                               else
                                 ""
                               end
              shape_line_xml = if shape[:no_line]
                                 "<a:ln><a:noFill/></a:ln>"
                               elsif shape[:line_color] || shape[:line_dash] || shape[:line_custom_dash] || shape[:head_end] || shape[:tail_end] || shape[:line_cap] || shape[:line_align] || shape[:line_compound] || shape[:line_join]
                                 ln_attrs = +(shape[:line_width] ? %( w="#{shape[:line_width].to_i}") : "")
                                 ln_attrs << %( cap="#{xml_escape(shape[:line_cap])}") if shape[:line_cap]
                                 ln_attrs << %( algn="#{xml_escape(shape[:line_align])}") if shape[:line_align]
                                 ln_attrs << %( cmpd="#{xml_escape(shape[:line_compound])}") if shape[:line_compound]
                                 fill_part = if shape[:line_color]
                                               "<a:solidFill>#{srgb_clr_xml(shape[:line_color], alpha: shape[:line_alpha], transforms: shape[:line_color_transforms])}</a:solidFill>"
                                             else
                                               ""
                                             end
                                 dash_part = if shape[:line_custom_dash]
                                               ds_xml = shape[:line_custom_dash].map { |ds| %(<a:ds d="#{ds[:d]}" sp="#{ds[:sp]}"/>) }.join
                                               "<a:custDash>#{ds_xml}</a:custDash>"
                                             elsif shape[:line_dash]
                                               %(<a:prstDash val="#{xml_escape(shape[:line_dash])}"/>)
                                             else
                                               ""
                                             end
                                 head_part = build_line_end_xml("headEnd", shape[:head_end])
                                 tail_part = build_line_end_xml("tailEnd", shape[:tail_end])
                                 join_part = case shape[:line_join]
                                             when "round" then "<a:round/>"
                                             when "bevel" then "<a:bevel/>"
                                             when "miter"
                                               shape[:line_miter_limit] ? "<a:miter lim=\"#{shape[:line_miter_limit].to_i}\"/>" : "<a:miter/>"
                                             else ""
                                             end
                                 %(<a:ln#{ln_attrs}>#{fill_part}#{dash_part}#{head_part}#{tail_part}#{join_part}</a:ln>)
                               else
                                 ""
                               end
              shape_xfrm_xml = shape[:rotation] ? %(<a:xfrm rot="#{shape[:rotation].to_i}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>) : ""
              av_lst_xml = if shape[:adjust_values]&.any?
                             gds = shape[:adjust_values].map { |gd| %(<a:gd name="#{xml_escape(gd[:name])}" fmla="#{xml_escape(gd[:fmla])}"/>) }
                             "<a:avLst>#{gds.join}</a:avLst>"
                           else
                             "<a:avLst/>"
                           end
              effect_lst_xml = if shape[:outer_shadow] || shape[:inner_shadow] || shape[:glow] || shape[:soft_edge] || shape[:reflection] || shape[:blur]
                                 effect_children = +""
                                 if shape[:outer_shadow]
                                   os = shape[:outer_shadow]
                                   os_attrs = +""
                                   os_attrs << %( blurRad="#{os[:blur_rad]}") if os[:blur_rad]
                                   os_attrs << %( dist="#{os[:dist]}") if os[:dist]
                                   os_attrs << %( dir="#{os[:dir]}") if os[:dir]
                                   os_attrs << %( algn="#{xml_escape(os[:algn])}") if os[:algn]
                                   os_attrs << %( rotWithShape="#{os[:rot_with_shape] ? 1 : 0}") unless os[:rot_with_shape].nil?
                                   os_color = os[:color] ? srgb_clr_xml(os[:color], alpha: os[:alpha]) : ""
                                   effect_children << "<a:outerShdw#{os_attrs}>#{os_color}</a:outerShdw>"
                                 end
                                 if shape[:glow]
                                   gl = shape[:glow]
                                   gl_attrs = +""
                                   gl_attrs << %( rad="#{gl[:rad]}") if gl[:rad]
                                   gl_color = gl[:color] ? srgb_clr_xml(gl[:color], alpha: gl[:alpha]) : ""
                                   effect_children << "<a:glow#{gl_attrs}>#{gl_color}</a:glow>"
                                 end
                                 if shape[:inner_shadow]
                                   is = shape[:inner_shadow]
                                   is_attrs = +""
                                   is_attrs << %( blurRad="#{is[:blur_rad]}") if is[:blur_rad]
                                   is_attrs << %( dist="#{is[:dist]}") if is[:dist]
                                   is_attrs << %( dir="#{is[:dir]}") if is[:dir]
                                   is_color = is[:color] ? srgb_clr_xml(is[:color], alpha: is[:alpha]) : ""
                                   effect_children << "<a:innerShdw#{is_attrs}>#{is_color}</a:innerShdw>"
                                 end
                                 if shape[:reflection]
                                   rf = shape[:reflection]
                                   rf_attrs = +""
                                   rf_attrs << %( blurRad="#{rf[:blur_rad]}") if rf[:blur_rad]
                                   rf_attrs << %( stA="#{rf[:st_a]}") if rf[:st_a]
                                   rf_attrs << %( endA="#{rf[:end_a]}") if rf[:end_a]
                                   rf_attrs << %( dist="#{rf[:dist]}") unless rf[:dist].nil?
                                   rf_attrs << %( dir="#{rf[:dir]}") if rf[:dir]
                                   rf_attrs << %( fadeDir="#{rf[:fade_dir]}") if rf[:fade_dir]
                                   rf_attrs << %( sx="#{rf[:sx]}") if rf[:sx]
                                   rf_attrs << %( sy="#{rf[:sy]}") if rf[:sy]
                                   rf_attrs << %( kx="#{rf[:kx]}") if rf[:kx]
                                   rf_attrs << %( ky="#{rf[:ky]}") if rf[:ky]
                                   rf_attrs << %( algn="#{xml_escape(rf[:algn])}") if rf[:algn]
                                   rf_attrs << %( rotWithShape="#{rf[:rot_with_shape] ? 1 : 0}") unless rf[:rot_with_shape].nil?
                                   effect_children << "<a:reflection#{rf_attrs}/>"
                                 end
                                 if shape[:soft_edge]
                                   se = shape[:soft_edge]
                                   effect_children << %(<a:softEdge rad="#{se[:rad]}"/>)
                                 end
                                 if shape[:blur]
                                   bl = shape[:blur]
                                   bl_attrs = +""
                                   bl_attrs << %( rad="#{bl[:rad]}") if bl[:rad]
                                   bl_attrs << %( grow="#{bl[:grow] ? 1 : 0}") unless bl[:grow].nil?
                                   effect_children << "<a:blur#{bl_attrs}/>"
                                 end
                                 "<a:effectLst>#{effect_children}</a:effectLst>"
                               else
                                 ""
                               end
              parts << %(<xdr:spPr>#{shape_xfrm_xml}<a:prstGeom prst="#{xml_escape(shape[:preset])}">#{av_lst_xml}</a:prstGeom>#{shape_fill_xml}#{shape_line_xml}#{effect_lst_xml}</xdr:spPr>)
              if shape[:text] || shape[:text_paragraphs]
                body_pr_attrs = +""
                body_pr_attrs << %( rot="#{shape[:text_rot]}") if shape[:text_rot]
                body_pr_attrs << %( spcFirstLastPara="1") if shape[:text_spc_first_last_para]
                body_pr_attrs << %( wrap="#{xml_escape(shape[:text_wrap])}") if shape[:text_wrap]
                body_pr_attrs << %( anchor="#{xml_escape(shape[:text_anchor])}") if shape[:text_anchor]
                body_pr_attrs << %( anchorCtr="1") if shape[:text_anchor_ctr]
                body_pr_attrs << %( vertOverflow="#{xml_escape(shape[:text_vert_overflow])}") if shape[:text_vert_overflow]
                body_pr_attrs << %( horzOverflow="#{xml_escape(shape[:text_horz_overflow])}") if shape[:text_horz_overflow]
                body_pr_attrs << %( numCol="#{shape[:text_num_col]}") if shape[:text_num_col]
                body_pr_attrs << %( spcCol="#{shape[:text_spc_col]}") if shape[:text_spc_col]
                body_pr_attrs << %( rtlCol="1") if shape[:text_rtl_col]
                body_pr_attrs << %( fromWordArt="1") if shape[:text_from_word_art]
                body_pr_attrs << %( upright="1") if shape[:text_upright]
                body_pr_attrs << %( compatLnSpc="1") if shape[:text_compat_ln_spc]
                body_pr_attrs << %( forceAA="1") if shape[:text_force_aa]
                body_pr_attrs << %( vert="#{xml_escape(shape[:text_vertical])}") if shape[:text_vertical]
                if shape[:text_insets]
                  ins = shape[:text_insets]
                  body_pr_attrs << %( lIns="#{ins[:left]}") if ins[:left]
                  body_pr_attrs << %( tIns="#{ins[:top]}") if ins[:top]
                  body_pr_attrs << %( rIns="#{ins[:right]}") if ins[:right]
                  body_pr_attrs << %( bIns="#{ins[:bottom]}") if ins[:bottom]
                end
                body_pr_children = +""
                if shape[:text_warp]
                  tw = shape[:text_warp]
                  body_pr_children << %(<a:prstTxWarp prst="#{xml_escape(tw[:preset])}"><a:avLst/></a:prstTxWarp>)
                end
                autofit_xml = case shape[:autofit]
                              when "none" then "<a:noAutofit/>"
                              when "shape" then "<a:spAutoFit/>"
                              when "normal" then "<a:normAutofit/>"
                              when Hash
                                af = shape[:autofit]
                                case af[:type]
                                when "normal"
                                  na_attrs = +""
                                  na_attrs << %( fontScale="#{af[:font_scale]}") if af[:font_scale]
                                  na_attrs << %( lnSpcReduction="#{af[:ln_spc_reduction]}") if af[:ln_spc_reduction]
                                  "<a:normAutofit#{na_attrs}/>"
                                when "none" then "<a:noAutofit/>"
                                when "shape" then "<a:spAutoFit/>"
                                else ""
                                end
                              else ""
                              end
                body_pr_children << autofit_xml
                body_pr_xml = if body_pr_children.empty?
                                "<a:bodyPr#{body_pr_attrs}/>"
                              else
                                "<a:bodyPr#{body_pr_attrs}>#{body_pr_children}</a:bodyPr>"
                              end
                parts << "<xdr:txBody>#{body_pr_xml}<a:lstStyle/>"
                if shape[:text_paragraphs]
                  shape[:text_paragraphs].each { |para| parts << paragraph_xml(para) }
                elsif shape[:text]
                  parts << paragraph_xml(
                    text: shape[:text], font: shape[:text_font],
                    align: shape[:text_align], font_align: shape[:text_font_align],
                    def_tab_sz: shape[:text_def_tab_sz], rtl: shape[:text_rtl],
                    ea_ln_brk: shape[:text_ea_ln_brk], latin_ln_brk: shape[:text_latin_ln_brk],
                    hanging_punct: shape[:text_hanging_punct], level: shape[:text_level],
                    indent: shape[:text_indent], spacing: shape[:text_spacing],
                    tab_stops: shape[:text_tab_stops], bullet: shape[:text_bullet],
                    def_rpr: shape[:text_def_rpr], end_para_rpr: shape[:text_end_para_rpr]
                  )
                end
                parts << "</xdr:txBody>"
              end
              parts << "</xdr:sp>"
              parts << client_data_xml(shape)
              parts << "</xdr:twoCellAnchor>"
            end
          end

          parts << "</xdr:wsDr>"
          parts.join
        end

        # : (untyped para) -> ::String
        def paragraph_xml(para)
          ppr_attrs = +""
          ppr_attrs << %( algn="#{xml_escape(para[:align])}") if para[:align]
          ppr_attrs << %( fontAlgn="#{xml_escape(para[:font_align])}") if para[:font_align]
          ppr_attrs << %( defTabSz="#{para[:def_tab_sz]}") if para[:def_tab_sz]
          ppr_attrs << %( rtl="1") if para[:rtl]
          ppr_attrs << %( eaLnBrk="1") if para[:ea_ln_brk]
          ppr_attrs << %( latinLnBrk="1") if para[:latin_ln_brk]
          ppr_attrs << %( hangingPunct="1") if para[:hanging_punct]
          ppr_attrs << %( lvl="#{para[:level]}") if para[:level]
          if para[:indent]
            ti = para[:indent]
            ppr_attrs << %( marL="#{ti[:left]}") if ti[:left]
            ppr_attrs << %( marR="#{ti[:right]}") if ti[:right]
            ppr_attrs << %( indent="#{ti[:indent]}") if ti[:indent]
          end
          ppr_children = +""
          if para[:spacing]
            ts = para[:spacing]
            if ts[:line]
              ppr_children << %(<a:lnSpc><a:spcPts val="#{ts[:line]}"/></a:lnSpc>)
            elsif ts[:line_pct]
              ppr_children << %(<a:lnSpc><a:spcPct val="#{ts[:line_pct]}"/></a:lnSpc>)
            end
            if ts[:before]
              ppr_children << %(<a:spcBef><a:spcPts val="#{ts[:before]}"/></a:spcBef>)
            elsif ts[:before_pct]
              ppr_children << %(<a:spcBef><a:spcPct val="#{ts[:before_pct]}"/></a:spcBef>)
            end
            if ts[:after]
              ppr_children << %(<a:spcAft><a:spcPts val="#{ts[:after]}"/></a:spcAft>)
            elsif ts[:after_pct]
              ppr_children << %(<a:spcAft><a:spcPct val="#{ts[:after_pct]}"/></a:spcAft>)
            end
          end
          if para[:tab_stops]&.any?
            tabs = para[:tab_stops].map do |tab|
              tab_attrs = %( pos="#{tab[:pos]}")
              tab_attrs << %( algn="#{xml_escape(tab[:align])}") if tab[:align]
              "<a:tab#{tab_attrs}/>"
            end
            ppr_children << "<a:tabLst>#{tabs.join}</a:tabLst>"
          end
          if para[:bullet]
            bu = para[:bullet]
            ppr_children << %(<a:buClr>#{color_xml(bu[:color])}</a:buClr>) if bu[:color]
            ppr_children << %(<a:buSzPts val="#{bu[:size_pts]}"/>) if bu[:size_pts]
            ppr_children << %(<a:buSzPct val="#{bu[:size_pct]}"/>) if bu[:size_pct]
            ppr_children << %(<a:buFont typeface="#{xml_escape(bu[:font])}"/>) if bu[:font]
            case bu[:type]
            when "none"
              ppr_children << "<a:buNone/>"
            when "char"
              ppr_children << %(<a:buChar char="#{xml_escape(bu[:char])}"/>)
            when "auto"
              auto_attrs = %( type="#{xml_escape(bu[:auto_type])}")
              auto_attrs << %( startAt="#{bu[:start_at]}") if bu[:start_at]
              ppr_children << "<a:buAutoNum#{auto_attrs}/>"
            end
          end
          ppr_children << text_char_props_xml("a:defRPr", para[:def_rpr]) if para[:def_rpr]
          ppr_xml = if ppr_attrs.empty? && ppr_children.empty?
                      ""
                    elsif ppr_children.empty?
                      "<a:pPr#{ppr_attrs}/>"
                    else
                      "<a:pPr#{ppr_attrs}>#{ppr_children}</a:pPr>"
                    end
          end_para_rpr_xml = para[:end_para_rpr] ? text_char_props_xml("a:endParaRPr", para[:end_para_rpr]) : ""
          runs_xml = if para[:runs]
                       para[:runs].map do |run|
                         run_rpr = run[:font] ? text_char_props_xml("a:rPr", run[:font]) : ""
                         "<a:r>#{run_rpr}<a:t>#{xml_escape(run[:text] || "")}</a:t></a:r>"
                       end.join
                     else
                       rpr_xml = para[:font] ? text_char_props_xml("a:rPr", para[:font]) : ""
                       "<a:r>#{rpr_xml}<a:t>#{xml_escape(para[:text] || "")}</a:t></a:r>"
                     end
          "<a:p>#{ppr_xml}#{runs_xml}#{end_para_rpr_xml}</a:p>"
        end

        # : (untyped tag, untyped font) -> untyped
        def text_char_props_xml(tag, font)
          attrs = +""
          attrs << %( b="1") if font[:bold]
          attrs << %( i="1") if font[:italic]
          attrs << %( noProof="1") if font[:no_proof]
          attrs << %( normalizeH="1") if font[:normalize_h]
          attrs << %( kumimoji="1") if font[:kumimoji]
          attrs << %( sz="#{font[:size]}") if font[:size]
          attrs << %( strike="#{xml_escape(font[:strike])}") if font[:strike]
          attrs << %( u="#{xml_escape(font[:underline])}") if font[:underline]
          attrs << %( baseline="#{font[:baseline]}") if font[:baseline]
          attrs << %( spc="#{font[:spacing]}") if font[:spacing]
          attrs << %( kern="#{font[:kern]}") if font[:kern]
          attrs << %( cap="#{xml_escape(font[:cap])}") if font[:cap]
          attrs << %( lang="#{xml_escape(font[:lang])}") if font[:lang]
          attrs << %( altLang="#{xml_escape(font[:alt_lang])}") if font[:alt_lang]
          attrs << %( dirty="1") if font[:dirty]
          attrs << %( smtClean="1") if font[:smt_clean]
          attrs << %( err="1") if font[:err]
          attrs << %( bmk="#{xml_escape(font[:bmk])}") if font[:bmk]
          children = +""
          children << %(<a:solidFill>#{color_xml(font[:color])}</a:solidFill>) if font[:color]
          children << %(<a:highlight>#{color_xml(font[:highlight])}</a:highlight>) if font[:highlight]
          children << %(<a:latin typeface="#{xml_escape(font[:name])}"/>) if font[:name]
          children << %(<a:ea typeface="#{xml_escape(font[:ea_font])}"/>) if font[:ea_font]
          children << %(<a:cs typeface="#{xml_escape(font[:cs_font])}"/>) if font[:cs_font]
          children << %(<a:sym typeface="#{xml_escape(font[:sym_font])}"/>) if font[:sym_font]
          children << "<a:uFillTx/>" if font[:u_fill_tx]
          children << "<a:uLnTx/>" if font[:u_ln_tx]
          if font[:line_color] || font[:line_width] || font[:line_dash] || font[:line_cap] || font[:line_join]
            ln_attrs = +(font[:line_width] ? %( w="#{font[:line_width].to_i}") : "")
            ln_attrs << %( cap="#{xml_escape(font[:line_cap])}") if font[:line_cap]
            ln_fill = font[:line_color] ? %(<a:solidFill>#{color_xml(font[:line_color])}</a:solidFill>) : ""
            ln_dash = font[:line_dash] ? %(<a:prstDash val="#{xml_escape(font[:line_dash])}"/>) : ""
            ln_join = case font[:line_join]
                      when "round" then "<a:round/>"
                      when "bevel" then "<a:bevel/>"
                      when "miter"
                        font[:line_miter_limit] ? "<a:miter lim=\"#{font[:line_miter_limit].to_i}\"/>" : "<a:miter/>"
                      else ""
                      end
            children << "<a:ln#{ln_attrs}>#{ln_fill}#{ln_dash}#{ln_join}</a:ln>"
          end
          if font[:outer_shadow] || font[:inner_shadow] || font[:glow] || font[:soft_edge] || font[:reflection] || font[:blur]
            eff = +""
            if font[:outer_shadow]
              os = font[:outer_shadow]
              oa = +""
              oa << %( blurRad="#{os[:blur_rad]}") if os[:blur_rad]
              oa << %( dist="#{os[:dist]}") if os[:dist]
              oa << %( dir="#{os[:dir]}") if os[:dir]
              oa << %( algn="#{xml_escape(os[:algn])}") if os[:algn]
              oa << %( rotWithShape="#{os[:rot_with_shape] ? 1 : 0}") unless os[:rot_with_shape].nil?
              oc = os[:color] ? srgb_clr_xml(os[:color], alpha: os[:alpha]) : ""
              eff << "<a:outerShdw#{oa}>#{oc}</a:outerShdw>"
            end
            if font[:glow]
              gl = font[:glow]
              ga = +(gl[:rad] ? %( rad="#{gl[:rad]}") : "")
              gc = gl[:color] ? srgb_clr_xml(gl[:color], alpha: gl[:alpha]) : ""
              eff << "<a:glow#{ga}>#{gc}</a:glow>"
            end
            if font[:inner_shadow]
              is = font[:inner_shadow]
              ia = +""
              ia << %( blurRad="#{is[:blur_rad]}") if is[:blur_rad]
              ia << %( dist="#{is[:dist]}") if is[:dist]
              ia << %( dir="#{is[:dir]}") if is[:dir]
              ic = is[:color] ? srgb_clr_xml(is[:color], alpha: is[:alpha]) : ""
              eff << "<a:innerShdw#{ia}>#{ic}</a:innerShdw>"
            end
            if font[:reflection]
              rf = font[:reflection]
              ra = +""
              ra << %( blurRad="#{rf[:blur_rad]}") if rf[:blur_rad]
              ra << %( stA="#{rf[:st_a]}") if rf[:st_a]
              ra << %( endA="#{rf[:end_a]}") if rf[:end_a]
              ra << %( dist="#{rf[:dist]}") unless rf[:dist].nil?
              ra << %( dir="#{rf[:dir]}") if rf[:dir]
              ra << %( fadeDir="#{rf[:fade_dir]}") if rf[:fade_dir]
              ra << %( sx="#{rf[:sx]}") if rf[:sx]
              ra << %( sy="#{rf[:sy]}") if rf[:sy]
              ra << %( kx="#{rf[:kx]}") if rf[:kx]
              ra << %( ky="#{rf[:ky]}") if rf[:ky]
              ra << %( algn="#{xml_escape(rf[:algn])}") if rf[:algn]
              ra << %( rotWithShape="#{rf[:rot_with_shape] ? 1 : 0}") unless rf[:rot_with_shape].nil?
              eff << "<a:reflection#{ra}/>"
            end
            if font[:soft_edge]
              se = font[:soft_edge]
              eff << %(<a:softEdge rad="#{se[:rad]}"/>)
            end
            if font[:blur]
              bl = font[:blur]
              ba = +(bl[:rad] ? %( rad="#{bl[:rad]}") : "")
              ba << %( grow="#{bl[:grow] ? 1 : 0}") unless bl[:grow].nil?
              eff << "<a:blur#{ba}/>"
            end
            children << "<a:effectLst>#{eff}</a:effectLst>"
          end
          if children.empty?
            "<#{tag}#{attrs}/>"
          else
            "<#{tag}#{attrs}>#{children}</#{tag}>"
          end
        end

        # : (untyped tag, untyped col, untyped row, ?col_off: ::Integer, ?row_off: ::Integer) -> ::String
        def anchor_xml(tag, col, row, col_off: 0, row_off: 0)
          "<xdr:#{tag}><xdr:col>#{col}</xdr:col><xdr:colOff>#{col_off}</xdr:colOff><xdr:row>#{row}</xdr:row><xdr:rowOff>#{row_off}</xdr:rowOff></xdr:#{tag}>"
        end

        # : (untyped obj) -> ::String
        def client_data_xml(obj)
          cd_attrs = +""
          cd_attrs << ' fLocksWithSheet="0"' if obj[:locks_with_sheet] == false
          cd_attrs << ' fLocksWithSheet="1"' if obj[:locks_with_sheet] == true
          cd_attrs << ' fPrintsWithSheet="0"' if obj[:prints_with_sheet] == false
          cd_attrs << ' fPrintsWithSheet="1"' if obj[:prints_with_sheet] == true
          "<xdr:clientData#{cd_attrs}/>"
        end

        # : (untyped rels_data) -> untyped
        def generate_drawing_rels(rels_data)
          parts = [XML_HEADER, %(<Relationships xmlns="#{REL_NS}">)]
          rels_data.each_with_index do |rel, i|
            rel_type = case rel[:type]
                       when :image then "#{DOC_REL_NS}/image"
                       when :chart then "#{DOC_REL_NS}/chart"
                       end
            parts << %(<Relationship Id="rId#{i + 1}" Type="#{rel_type}" Target="#{rel[:target]}"/>)
          end
          parts << "</Relationships>"
          parts.join
        end

        # : (untyped title_spec, ?overlay: bool) -> untyped
        def build_chart_title_xml(title_spec, overlay: false)
          overlay_val = overlay ? 1 : 0
          if title_spec.is_a?(Hash)
            text = xml_escape(title_spec[:text].to_s)
            rotation = title_spec[:rotation]
            body_pr = rotation ? %(<a:bodyPr rot="#{rotation}"/>) : "<a:bodyPr/>"
            font = title_spec[:font] || {}
            rpr_attrs = +""
            rpr_attrs << %( b="1") if font[:bold]
            rpr_attrs << %( i="1") if font[:italic]
            rpr_attrs << %( sz="#{font[:size]}") if font[:size]
            rpr_children = +""
            rpr_children << %(<a:solidFill>#{color_xml(font[:color])}</a:solidFill>) if font[:color]
            rpr_children << %(<a:latin typeface="#{xml_escape(font[:name])}"/>) if font[:name]
            rpr_xml = if rpr_children.empty?
                        "<a:rPr#{rpr_attrs}/>"
                      else
                        "<a:rPr#{rpr_attrs}>#{rpr_children}</a:rPr>"
                      end
            layout_xml = build_title_layout_xml(title_spec[:layout])
            sp_xml = build_title_sp_pr(title_spec)
            "<c:title><c:tx><c:rich>#{body_pr}<a:lstStyle/><a:p><a:r>#{rpr_xml}<a:t>#{text}</a:t></a:r></a:p></c:rich></c:tx>#{layout_xml}<c:overlay val=\"#{overlay_val}\"/>#{sp_xml}</c:title>"
          else
            "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(title_spec)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"#{overlay_val}\"/></c:title>"
          end
        end

        # : (untyped title_spec, untyped chart, untyped prefix) -> untyped
        def merge_flat_title_styling(title_spec, chart, prefix)
          font_key = :"#{prefix}_font"
          fill_key = :"#{prefix}_fill"
          no_fill_key = :"#{prefix}_no_fill"
          lc_key = :"#{prefix}_line_color"
          lw_key = :"#{prefix}_line_width"
          ld_key = :"#{prefix}_line_dash"
          rot_key = :"#{prefix}_rotation"
          return title_spec unless chart[font_key] || chart[fill_key] || chart[no_fill_key] ||
                                   chart[lc_key] || chart[lw_key] || chart[ld_key] || chart[rot_key]

          spec = title_spec.is_a?(Hash) ? title_spec : { text: title_spec }
          spec[:font] ||= chart[font_key] if chart[font_key]
          spec[:fill_color] ||= chart[fill_key] if chart[fill_key]
          spec[:no_fill] = chart[no_fill_key] if spec[:no_fill].nil? && !chart[no_fill_key].nil?
          spec[:line_color] ||= chart[lc_key] if chart[lc_key]
          spec[:line_width] ||= chart[lw_key] if chart[lw_key]
          spec[:line_dash] ||= chart[ld_key] if chart[ld_key]
          spec[:rotation] ||= chart[rot_key] if chart[rot_key]
          spec
        end

        # : (untyped spec) -> untyped
        def build_title_sp_pr(spec)
          children = +""
          children << %(<a:solidFill>#{color_xml(spec[:fill_color])}</a:solidFill>) if spec[:fill_color]
          children << "<a:noFill/>" if spec[:no_fill]
          if spec[:line_color] || spec[:line_width] || spec[:line_dash]
            lw = spec[:line_width] ? %( w="#{(spec[:line_width] * 12_700).to_i}") : ""
            lf = spec[:line_color] ? %(<a:solidFill>#{color_xml(spec[:line_color])}</a:solidFill>) : ""
            ld = spec[:line_dash] ? %(<a:prstDash val="#{xml_escape(spec[:line_dash])}"/>) : ""
            children << "<a:ln#{lw}>#{lf}#{ld}</a:ln>"
          end
          children.empty? ? "" : "<c:spPr>#{children}</c:spPr>"
        end

        # : (untyped layout) -> ("" | untyped)
        def build_title_layout_xml(layout)
          return "" unless layout.is_a?(Hash)

          ml = +""
          ml << %(<c:x val="#{layout[:x]}"/>) if layout[:x]
          ml << %(<c:y val="#{layout[:y]}"/>) if layout[:y]
          ml << %(<c:w val="#{layout[:w]}"/>) if layout[:w]
          ml << %(<c:h val="#{layout[:h]}"/>) if layout[:h]
          ml.empty? ? "" : "<c:layout><c:manualLayout>#{ml}</c:manualLayout></c:layout>"
        end

        # : (untyped tag, untyped spec) -> ("" | untyped)
        def gridlines_xml(tag, spec)
          return "" unless spec

          if spec.is_a?(Hash)
            sp_children = +""
            if spec[:line_color] || spec[:line_width] || spec[:line_dash]
              lw = spec[:line_width] ? %( w="#{(spec[:line_width] * 12_700).to_i}") : ""
              lf = spec[:line_color] ? %(<a:solidFill>#{color_xml(spec[:line_color])}</a:solidFill>) : ""
              ld = spec[:line_dash] ? %(<a:prstDash val="#{xml_escape(spec[:line_dash])}"/>) : ""
              sp_children << "<a:ln#{lw}>#{lf}#{ld}</a:ln>"
            end
            sp_children.empty? ? "<c:#{tag}/>" : "<c:#{tag}><c:spPr>#{sp_children}</c:spPr></c:#{tag}>"
          else
            "<c:#{tag}/>"
          end
        end

        # Resolves a cell reference like "Sheet1!$A$1:$A$5" to an array of values from @sheets.
        # : (untyped ref) -> (nil | untyped)
        def resolve_sheet_ref(ref)
          return nil unless ref

          m = if ref.start_with?("'")
                ref.match(/\A'([^']+)'!(.+)\z/)

              else
                ref.match(/\A([^!]+)!(.+)\z/)

              end
          return nil unless m

          sheet_name = m[1]
          range_part = m[2]
          return nil unless @sheets.key?(sheet_name)

          range_part = range_part.delete("$")
          cells = if range_part.include?(":")
                    start_cell, end_cell = range_part.split(":", 2)
                    enumerate_cell_range(start_cell, end_cell)
                  else
                    [range_part]
                  end
          cells.map { |addr| @sheets[sheet_name][addr] }
        end

        # : (untyped start_cell, untyped end_cell) -> (::Array[untyped] | untyped)
        def enumerate_cell_range(start_cell, end_cell)
          sc = start_cell.match(/\A([A-Z]+)(\d+)\z/)
          ec = end_cell.match(/\A([A-Z]+)(\d+)\z/)
          return [start_cell] unless sc && ec

          start_col = column_letter_to_index(sc[1])
          start_row = sc[2].to_i
          end_col = column_letter_to_index(ec[1])
          end_row = ec[2].to_i
          cells = []
          (start_row..end_row).each do |row|
            (start_col..end_col).each do |col|
              cells << "#{index_to_column_letter(col)}#{row}"
            end
          end
          cells
        end

        # : (untyped ref) -> ("" | ::String)
        def num_cache_xml(ref)
          values = resolve_sheet_ref(ref)
          return "" unless values

          pts = values.each_with_index.filter_map do |v, i|
            next unless v

            %(<c:pt idx="#{i}"><c:v>#{xml_escape(v.to_s)}</c:v></c:pt>)
          end
          %(<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="#{values.size}"/>#{pts.join}</c:numCache>)
        end

        # : (untyped ref) -> ("" | ::String)
        def str_cache_xml(ref)
          values = resolve_sheet_ref(ref)
          return "" unless values

          pts = values.each_with_index.filter_map do |v, i|
            next unless v

            %(<c:pt idx="#{i}"><c:v>#{xml_escape(v.to_s)}</c:v></c:pt>)
          end
          %(<c:strCache><c:ptCount val="#{values.size}"/>#{pts.join}</c:strCache>)
        end

        # : (untyped color, ?alpha: untyped?, ?transforms: untyped?) -> untyped
        def color_xml(color, alpha: nil, transforms: nil)
          if color.is_a?(Hash) && color[:scheme]
            tag = "schemeClr"
            val = xml_escape(color[:scheme])
            transforms = color[:transforms] || transforms
          else
            tag = "srgbClr"
            val = xml_escape(normalize_drawing_rgb(color))
          end
          children = +""
          children << %(<a:alpha val="#{alpha}"/>) if alpha
          (transforms || []).each do |t|
            children << %(<a:#{xml_escape(t[:type])} val="#{t[:val]}"/>)
          end
          if children.empty?
            %(<a:#{tag} val="#{val}"/>)
          else
            %(<a:#{tag} val="#{val}">#{children}</a:#{tag}>)
          end
        end

        # : (untyped color) -> untyped
        def normalize_drawing_rgb(color)
          str = color.to_s
          hex = str.delete_prefix("#")

          # Common shorthand (e.g. #f00 -> FF0000)
          return hex.chars.map { |c| c * 2 }.join.upcase if hex.match?(/\A[0-9A-Fa-f]{3}\z/)

          # ARGB -> RGB for DrawingML srgbClr (expects RRGGBB)
          return hex[-6, 6].upcase if hex.match?(/\A[0-9A-Fa-f]{8}\z/)

          return hex.upcase if hex.match?(/\A[0-9A-Fa-f]{6}\z/)

          str
        end

        alias srgb_clr_xml color_xml

        # : (untyped rotation, untyped font) -> ("" | ::String)
        def build_axis_txpr(rotation, font)
          return "" unless rotation || font

          body_attrs = rotation ? %( rot="#{rotation}") : ""
          rpr_attrs = +""
          rpr_children = +""
          if font
            rpr_attrs << %( sz="#{(font[:size] * 100).to_i}") if font[:size]
            rpr_attrs << %( b="1") if font[:bold]
            rpr_attrs << %( i="1") if font[:italic]
            rpr_children << %(<a:solidFill>#{color_xml(font[:color])}</a:solidFill>) if font[:color]
            rpr_children << %(<a:latin typeface="#{xml_escape(font[:name])}"/>) if font[:name]
          end
          def_rpr = if rpr_children.empty?
                      "<a:defRPr#{rpr_attrs}/>"
                    else
                      "<a:defRPr#{rpr_attrs}>#{rpr_children}</a:defRPr>"
                    end
          %(<c:txPr><a:bodyPr#{body_attrs}/><a:lstStyle/><a:p><a:pPr>#{def_rpr}</a:pPr><a:endParaRPr/></a:p></c:txPr>)
        end

        # : (untyped chart) -> untyped
        def generate_chart_xml(chart)
          chart_type = CHART_TYPE_MAP[chart[:type]] || "barChart"
          no_axes = NO_AXIS_CHARTS.include?(chart_type)
          parts = [
            XML_HEADER,
            %(<c:chartSpace xmlns:c="#{C_NS}" xmlns:a="#{A_NS}" xmlns:r="#{DOC_REL_NS}">)
          ]
          rc = chart[:rounded_corners]
          parts << %(<c:roundedCorners val="#{rc ? 1 : 0}"/>) unless rc.nil?
          parts << %(<c:style val="#{chart[:style]}"/>) if chart[:style]
          if (prot = chart[:protection])
            parts << "<c:protection>"
            parts << %(<c:chartObject val="#{prot[:chart_object] ? 1 : 0}"/>) unless prot[:chart_object].nil?
            parts << %(<c:data val="#{prot[:data] ? 1 : 0}"/>) unless prot[:data].nil?
            parts << %(<c:formatting val="#{prot[:formatting] ? 1 : 0}"/>) unless prot[:formatting].nil?
            parts << %(<c:selection val="#{prot[:selection] ? 1 : 0}"/>) unless prot[:selection].nil?
            parts << %(<c:userInterface val="#{prot[:user_interface] ? 1 : 0}"/>) unless prot[:user_interface].nil?
            parts << "</c:protection>"
          end
          parts << "<c:chart>"

          if chart[:title]
            title_spec = chart[:title]
            if chart[:title_font] || chart[:title_fill_color] || chart[:title_no_fill] ||
               chart[:title_line_color] || chart[:title_line_width] || chart[:title_line_dash] ||
               chart[:title_rotation]
              title_spec = { text: title_spec } unless title_spec.is_a?(Hash)
              title_spec[:font] ||= chart[:title_font] if chart[:title_font]
              title_spec[:fill_color] ||= chart[:title_fill_color] if chart[:title_fill_color]
              title_spec[:no_fill] = chart[:title_no_fill] if title_spec[:no_fill].nil? && !chart[:title_no_fill].nil?
              title_spec[:line_color] ||= chart[:title_line_color] if chart[:title_line_color]
              title_spec[:line_width] ||= chart[:title_line_width] if chart[:title_line_width]
              title_spec[:line_dash] ||= chart[:title_line_dash] if chart[:title_line_dash]
              title_spec[:rotation] ||= chart[:title_rotation] if chart[:title_rotation]
            end
            parts << build_chart_title_xml(title_spec, overlay: chart[:title_overlay])
          end
          atd = chart[:auto_title_deleted]
          parts << %(<c:autoTitleDeleted val="#{atd ? 1 : 0}"/>) unless atd.nil?

          if (v3d = chart[:view_3d])
            parts << "<c:view3D>"
            parts << %(<c:rotX val="#{v3d[:rot_x]}"/>) if v3d[:rot_x]
            parts << %(<c:hPercent val="#{v3d[:h_percent]}"/>) if v3d[:h_percent]
            parts << %(<c:rotY val="#{v3d[:rot_y]}"/>) if v3d[:rot_y]
            parts << %(<c:depthPercent val="#{v3d[:depth_percent]}"/>) if v3d[:depth_percent]
            r_ang = v3d[:r_ang_ax]
            parts << %(<c:rAngAx val="#{r_ang ? 1 : 0}"/>) unless r_ang.nil?
            parts << %(<c:perspective val="#{v3d[:perspective]}"/>) if v3d[:perspective]
            parts << "</c:view3D>"
          end

          %i[floor side_wall back_wall].each do |wall_key|
            next unless (wall = chart[wall_key])

            tag = { floor: "floor", side_wall: "sideWall", back_wall: "backWall" }[wall_key]
            parts << "<c:#{tag}><c:spPr>"
            parts << "<a:solidFill>#{srgb_clr_xml(wall[:fill_color])}</a:solidFill>" if wall[:fill_color]
            parts << "<a:noFill/>" if wall[:no_fill]
            if wall[:line_color] || wall[:line_dash]
              w_attr = wall[:line_width] ? %( w="#{wall[:line_width]}") : ""
              w_fill = wall[:line_color] ? %(<a:solidFill>#{srgb_clr_xml(wall[:line_color])}</a:solidFill>) : ""
              w_dash = wall[:line_dash] ? %(<a:prstDash val="#{xml_escape(wall[:line_dash])}"/>) : ""
              parts << "<a:ln#{w_attr}>#{w_fill}#{w_dash}</a:ln>"
            end
            parts << "</c:spPr></c:#{tag}>"
          end

          parts << "<c:plotArea>"
          pa_layout = chart[:plot_area_layout]
          if pa_layout.is_a?(Hash)
            ml_parts = +""
            ml_parts << %(<c:layoutTarget val="#{xml_escape(pa_layout[:target])}"/>) if pa_layout[:target]
            ml_parts << %(<c:xMode val="#{xml_escape(pa_layout[:x_mode])}"/>) if pa_layout[:x_mode]
            ml_parts << %(<c:yMode val="#{xml_escape(pa_layout[:y_mode])}"/>) if pa_layout[:y_mode]
            ml_parts << %(<c:wMode val="#{xml_escape(pa_layout[:w_mode])}"/>) if pa_layout[:w_mode]
            ml_parts << %(<c:hMode val="#{xml_escape(pa_layout[:h_mode])}"/>) if pa_layout[:h_mode]
            ml_parts << %(<c:x val="#{pa_layout[:x]}"/>) if pa_layout[:x]
            ml_parts << %(<c:y val="#{pa_layout[:y]}"/>) if pa_layout[:y]
            ml_parts << %(<c:w val="#{pa_layout[:w]}"/>) if pa_layout[:w]
            ml_parts << %(<c:h val="#{pa_layout[:h]}"/>) if pa_layout[:h]
            parts << if ml_parts.empty?
                       "<c:layout/>"
                     else
                       "<c:layout><c:manualLayout>#{ml_parts}</c:manualLayout></c:layout>"
                     end
          else
            parts << "<c:layout/>"
          end
          parts << "<c:#{chart_type}>"
          parts << %(<c:ofPieType val="#{xml_escape(chart[:of_pie_type] || "pie")}"/>) if chart_type == "ofPieChart"
          if %w[barChart bar3DChart].include?(chart_type)
            bd = chart[:bar_dir] || "col"
            gr = chart[:grouping] || "clustered"
            parts << %(<c:barDir val="#{bd}"/><c:grouping val="#{gr}"/>)
          elsif GROUPING_CHARTS.include?(chart_type)
            parts << %(<c:grouping val="#{chart[:grouping] || "standard"}"/>)
          end
          parts << %(<c:scatterStyle val="#{chart[:scatter_style] || "lineMarker"}"/>) if chart_type == "scatterChart"
          parts << %(<c:radarStyle val="#{chart[:radar_style] || "standard"}"/>) if chart_type == "radarChart"
          vc = chart[:vary_colors]
          parts << %(<c:varyColors val="#{vc ? 1 : 0}"/>) unless vc.nil?
          wf = chart[:wireframe]
          parts << %(<c:wireframe val="#{wf ? 1 : 0}"/>) unless wf.nil?

          all_series = chart[:series] || []
          all_series.each_with_index do |ser, idx|
            ser_order = ser[:order] || idx
            parts << "<c:ser><c:idx val=\"#{idx}\"/><c:order val=\"#{ser_order}\"/>"
            parts << "<c:tx><c:strRef><c:f>#{xml_escape(ser[:name])}</c:f>#{str_cache_xml(ser[:name])}</c:strRef></c:tx>" if ser[:name]
            iin = ser[:invert_if_negative]
            parts << %(<c:invertIfNegative val="#{iin ? 1 : 0}"/>) unless iin.nil?
            parts << %(<c:explosion val="#{ser[:explosion]}"/>) if ser[:explosion]
            ser[:data_points]&.each do |dp|
              parts << "<c:dPt><c:idx val=\"#{dp[:idx]}\"/>"
              dp_iin = dp[:invert_if_negative]
              parts << %(<c:invertIfNegative val="#{dp_iin ? 1 : 0}"/>) unless dp_iin.nil?
              if dp[:marker_symbol] || dp[:marker_size] || dp[:marker_fill] || dp[:marker_no_fill] || dp[:marker_line_color] || dp[:marker_line_dash] || dp[:marker_no_line]
                parts << "<c:marker>"
                parts << %(<c:symbol val="#{xml_escape(dp[:marker_symbol])}"/>) if dp[:marker_symbol]
                parts << %(<c:size val="#{dp[:marker_size]}"/>) if dp[:marker_size]
                if dp[:marker_fill] || dp[:marker_no_fill] || dp[:marker_line_color] || dp[:marker_line_dash] || dp[:marker_no_line]
                  mk_sp = +""
                  mk_sp << "<a:noFill/>" if dp[:marker_no_fill]
                  mk_sp << %(<a:solidFill>#{color_xml(dp[:marker_fill])}</a:solidFill>) if dp[:marker_fill]
                  if dp[:marker_no_line]
                    mk_sp << "<a:ln><a:noFill/></a:ln>"
                  elsif dp[:marker_line_color] || dp[:marker_line_dash]
                    mk_ln_w = dp[:marker_line_width] ? %( w="#{(dp[:marker_line_width] * 12_700).to_i}") : ""
                    mk_ln_f = dp[:marker_line_color] ? %(<a:solidFill>#{color_xml(dp[:marker_line_color])}</a:solidFill>) : ""
                    mk_ln_d = dp[:marker_line_dash] ? %(<a:prstDash val="#{xml_escape(dp[:marker_line_dash])}"/>) : ""
                    mk_sp << "<a:ln#{mk_ln_w}>#{mk_ln_f}#{mk_ln_d}</a:ln>"
                  end
                  parts << "<c:spPr>#{mk_sp}</c:spPr>"
                end
                parts << "</c:marker>"
              end
              dp_b3d = dp[:bubble_3d]
              parts << %(<c:bubble3D val="#{dp_b3d ? 1 : 0}"/>) unless dp_b3d.nil?
              parts << %(<c:explosion val="#{dp[:explosion]}"/>) if dp[:explosion]
              dp_sp_children = +""
              dp_sp_children << %(<a:solidFill>#{color_xml(dp[:fill_color])}</a:solidFill>) if dp[:fill_color]
              dp_sp_children << "<a:noFill/>" if dp[:no_fill]
              if dp[:no_line]
                dp_sp_children << "<a:ln><a:noFill/></a:ln>"
              elsif dp[:line_color] || dp[:line_width] || dp[:line_dash]
                dp_ln_attrs = dp[:line_width] ? %( w="#{(dp[:line_width] * 12_700).to_i}") : ""
                dp_ln_fill = dp[:line_color] ? %(<a:solidFill>#{color_xml(dp[:line_color])}</a:solidFill>) : ""
                dp_ln_dash = dp[:line_dash] ? %(<a:prstDash val="#{xml_escape(dp[:line_dash])}"/>) : ""
                dp_sp_children << "<a:ln#{dp_ln_attrs}>#{dp_ln_fill}#{dp_ln_dash}</a:ln>"
              end
              parts << "<c:spPr>#{dp_sp_children}</c:spPr>" unless dp_sp_children.empty?
              parts << "</c:dPt>"
            end
            if ser[:fill_color] || ser[:no_fill] || ser[:line_color] || ser[:no_line] || ser[:line_width] || ser[:line_cap] || ser[:line_join] || ser[:line_dash]
              parts << "<c:spPr>"
              parts << %(<a:solidFill>#{color_xml(ser[:fill_color])}</a:solidFill>) if ser[:fill_color]
              parts << "<a:noFill/>" if ser[:no_fill]
              if ser[:line_color] || ser[:no_line] || ser[:line_width] || ser[:line_cap] || ser[:line_join] || ser[:line_dash]
                lw = ser[:line_width] ? %( w="#{(ser[:line_width] * 12_700).to_i}") : ""
                lc = ser[:line_cap] ? %( cap="#{xml_escape(ser[:line_cap])}") : ""
                parts << "<a:ln#{lw}#{lc}>"
                parts << "<a:noFill/>" if ser[:no_line]
                parts << %(<a:solidFill>#{color_xml(ser[:line_color])}</a:solidFill>) if ser[:line_color]
                parts << %(<a:prstDash val="#{xml_escape(ser[:line_dash])}"/>) if ser[:line_dash]
                case ser[:line_join]
                when "round" then parts << "<a:round/>"
                when "bevel" then parts << "<a:bevel/>"
                when "miter"
                  parts << (ser[:line_miter_limit] ? "<a:miter lim=\"#{ser[:line_miter_limit].to_i}\"/>" : "<a:miter/>")
                end
                parts << "</a:ln>"
              end
              parts << "</c:spPr>"
            end
            if ser[:marker_symbol] || ser[:marker_size] || ser[:marker_fill] || ser[:marker_no_fill] || ser[:marker_line_color] || ser[:marker_line_dash] || ser[:marker_no_line]
              parts << "<c:marker>"
              parts << %(<c:symbol val="#{xml_escape(ser[:marker_symbol])}"/>) if ser[:marker_symbol]
              parts << %(<c:size val="#{ser[:marker_size]}"/>) if ser[:marker_size]
              if ser[:marker_fill] || ser[:marker_no_fill] || ser[:marker_line_color] || ser[:marker_line_dash] || ser[:marker_no_line]
                parts << "<c:spPr>"
                parts << "<a:noFill/>" if ser[:marker_no_fill]
                parts << %(<a:solidFill>#{color_xml(ser[:marker_fill])}</a:solidFill>) if ser[:marker_fill]
                if ser[:marker_no_line]
                  parts << "<a:ln><a:noFill/></a:ln>"
                elsif ser[:marker_line_color] || ser[:marker_line_dash]
                  mk_ln_w = ser[:marker_line_width] ? %( w="#{(ser[:marker_line_width] * 12_700).to_i}") : ""
                  mk_ln_f = ser[:marker_line_color] ? %(<a:solidFill>#{color_xml(ser[:marker_line_color])}</a:solidFill>) : ""
                  mk_ln_d = ser[:marker_line_dash] ? %(<a:prstDash val="#{xml_escape(ser[:marker_line_dash])}"/>) : ""
                  parts << "<a:ln#{mk_ln_w}>#{mk_ln_f}#{mk_ln_d}</a:ln>"
                end
                parts << "</c:spPr>"
              end
              parts << "</c:marker>"
            end
            if (dl = ser[:data_labels] || chart[:data_labels])
              parts << "<c:dLbls>"
              dl[:labels]&.each do |lbl|
                parts << "<c:dLbl>"
                parts << %(<c:idx val="#{lbl[:idx]}"/>)
                if lbl[:delete]
                  parts << '<c:delete val="1"/>'
                else
                  if (lbl_layout = lbl[:layout])
                    ml = +""
                    ml << %(<c:x val="#{lbl_layout[:x]}"/>) if lbl_layout[:x]
                    ml << %(<c:y val="#{lbl_layout[:y]}"/>) if lbl_layout[:y]
                    parts << (ml.empty? ? "<c:layout/>" : "<c:layout><c:manualLayout>#{ml}</c:manualLayout></c:layout>")
                  end
                  parts << "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(lbl[:text])}</a:t></a:r></a:p></c:rich></c:tx>" if lbl[:text]
                  if lbl[:num_fmt]
                    lnf = lbl[:num_fmt]
                    lnf_src = lnf[:source_linked] ? ' sourceLinked="1"' : ""
                    parts << %(<c:numFmt formatCode="#{xml_escape(lnf[:format_code])}"#{lnf_src}/>)
                  end
                  lbl_sp = +""
                  lbl_sp << %(<a:solidFill>#{color_xml(lbl[:fill_color])}</a:solidFill>) if lbl[:fill_color]
                  lbl_sp << "<a:noFill/>" if lbl[:no_fill]
                  if lbl[:line_color] || lbl[:line_width] || lbl[:line_dash]
                    lbl_lw = lbl[:line_width] ? %( w="#{(lbl[:line_width] * 12_700).to_i}") : ""
                    lbl_lf = lbl[:line_color] ? %(<a:solidFill>#{color_xml(lbl[:line_color])}</a:solidFill>) : ""
                    lbl_ld = lbl[:line_dash] ? %(<a:prstDash val="#{xml_escape(lbl[:line_dash])}"/>) : ""
                    lbl_sp << "<a:ln#{lbl_lw}>#{lbl_lf}#{lbl_ld}</a:ln>"
                  end
                  parts << "<c:spPr>#{lbl_sp}</c:spPr>" unless lbl_sp.empty?
                  parts << build_axis_txpr(nil, lbl[:font]) if lbl[:font]
                  parts << %(<c:dLblPos val="#{lbl[:position]}"/>) if lbl[:position]
                  parts << "<c:showLegendKey val=\"#{lbl[:show_legend_key] ? 1 : 0}\"/>" unless lbl[:show_legend_key].nil?
                  parts << "<c:showVal val=\"#{lbl[:show_val] ? 1 : 0}\"/>" unless lbl[:show_val].nil?
                  parts << "<c:showCatName val=\"#{lbl[:show_cat_name] ? 1 : 0}\"/>" unless lbl[:show_cat_name].nil?
                  parts << "<c:showSerName val=\"#{lbl[:show_ser_name] ? 1 : 0}\"/>" unless lbl[:show_ser_name].nil?
                  parts << "<c:showPercent val=\"#{lbl[:show_percent] ? 1 : 0}\"/>" unless lbl[:show_percent].nil?
                  parts << "<c:showBubbleSize val=\"#{lbl[:show_bubble_size] ? 1 : 0}\"/>" unless lbl[:show_bubble_size].nil?
                  parts << "<c:separator>#{xml_escape(lbl[:separator])}</c:separator>" if lbl[:separator]
                end
                parts << "</c:dLbl>"
              end
              if dl[:num_fmt]
                nf = dl[:num_fmt]
                nf_src = nf[:source_linked] ? ' sourceLinked="1"' : ""
                parts << %(<c:numFmt formatCode="#{xml_escape(nf[:format_code])}"#{nf_src}/>)
              end
              dl_sp_children = +""
              dl_sp_children << %(<a:solidFill>#{color_xml(dl[:fill_color])}</a:solidFill>) if dl[:fill_color]
              dl_sp_children << "<a:noFill/>" if dl[:no_fill]
              if dl[:line_color] || dl[:line_width] || dl[:line_dash]
                dl_lw = dl[:line_width] ? %( w="#{(dl[:line_width] * 12_700).to_i}") : ""
                dl_lf = dl[:line_color] ? %(<a:solidFill>#{color_xml(dl[:line_color])}</a:solidFill>) : ""
                dl_ld = dl[:line_dash] ? %(<a:prstDash val="#{xml_escape(dl[:line_dash])}"/>) : ""
                dl_sp_children << "<a:ln#{dl_lw}>#{dl_lf}#{dl_ld}</a:ln>"
              end
              parts << "<c:spPr>#{dl_sp_children}</c:spPr>" unless dl_sp_children.empty?
              parts << build_axis_txpr(nil, dl[:font]) if dl[:font]
              parts << %(<c:dLblPos val="#{dl[:position]}"/>) if dl[:position]
              parts << "<c:showLegendKey val=\"#{dl[:show_legend_key] ? 1 : 0}\"/>" unless dl[:show_legend_key].nil?
              parts << "<c:showVal val=\"#{dl[:show_val] ? 1 : 0}\"/>" unless dl[:show_val].nil?
              parts << "<c:showCatName val=\"#{dl[:show_cat_name] ? 1 : 0}\"/>" unless dl[:show_cat_name].nil?
              parts << "<c:showSerName val=\"#{dl[:show_ser_name] ? 1 : 0}\"/>" unless dl[:show_ser_name].nil?
              parts << "<c:showPercent val=\"#{dl[:show_percent] ? 1 : 0}\"/>" unless dl[:show_percent].nil?
              parts << "<c:showBubbleSize val=\"#{dl[:show_bubble_size] ? 1 : 0}\"/>" unless dl[:show_bubble_size].nil?
              parts << "<c:separator>#{xml_escape(dl[:separator])}</c:separator>" if dl[:separator]
              sl = dl[:show_leader_lines]
              parts << "<c:showLeaderLines val=\"#{sl ? 1 : 0}\"/>" unless sl.nil?
              if dl[:leader_lines]
                ll_spec = dl[:leader_lines]
                if ll_spec.is_a?(Hash)
                  ll_sp = +""
                  if ll_spec[:line_color] || ll_spec[:line_width] || ll_spec[:line_dash]
                    ll_lw = ll_spec[:line_width] ? %( w="#{(ll_spec[:line_width] * 12_700).to_i}") : ""
                    ll_lf = ll_spec[:line_color] ? %(<a:solidFill>#{color_xml(ll_spec[:line_color])}</a:solidFill>) : ""
                    ll_ld = ll_spec[:line_dash] ? %(<a:prstDash val="#{xml_escape(ll_spec[:line_dash])}"/>) : ""
                    ll_sp << "<a:ln#{ll_lw}>#{ll_lf}#{ll_ld}</a:ln>"
                  end
                  parts << (ll_sp.empty? ? "<c:leaderLines/>" : "<c:leaderLines><c:spPr>#{ll_sp}</c:spPr></c:leaderLines>")
                else
                  parts << "<c:leaderLines/>"
                end
              end
              parts << "</c:dLbls>"
            end
            trendline_list = ser[:trendlines] || (ser[:trendline] ? [ser[:trendline]] : [])
            trendline_list.each do |tl|
              parts << "<c:trendline>"
              parts << "<c:name>#{xml_escape(tl[:name])}</c:name>" if tl[:name]
              tl_sp_children = +""
              if tl[:line_color] || tl[:line_width] || tl[:line_dash]
                tl_ln_w = tl[:line_width] ? %( w="#{(tl[:line_width] * 12_700).to_i}") : ""
                tl_ln_f = tl[:line_color] ? %(<a:solidFill>#{color_xml(tl[:line_color])}</a:solidFill>) : ""
                tl_ln_d = tl[:line_dash] ? %(<a:prstDash val="#{xml_escape(tl[:line_dash])}"/>) : ""
                tl_sp_children << "<a:ln#{tl_ln_w}>#{tl_ln_f}#{tl_ln_d}</a:ln>"
              end
              parts << "<c:spPr>#{tl_sp_children}</c:spPr>" unless tl_sp_children.empty?
              parts << %(<c:trendlineType val="#{xml_escape(tl[:type] || "linear")}"/>)
              parts << %(<c:order val="#{tl[:order]}"/>) if tl[:order]
              parts << %(<c:period val="#{tl[:period]}"/>) if tl[:period]
              parts << %(<c:forward val="#{tl[:forward]}"/>) if tl[:forward]
              parts << %(<c:backward val="#{tl[:backward]}"/>) if tl[:backward]
              parts << %(<c:intercept val="#{tl[:intercept]}"/>) if tl[:intercept]
              parts << %(<c:dispRSqr val="#{tl[:disp_r_sqr] ? 1 : 0}"/>) unless tl[:disp_r_sqr].nil?
              parts << %(<c:dispEq val="#{tl[:disp_eq] ? 1 : 0}"/>) unless tl[:disp_eq].nil?
              if tl[:label]
                lbl = tl[:label]
                parts << "<c:trendlineLbl>"
                parts << build_title_layout_xml(lbl[:layout])
                parts << "<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>#{xml_escape(lbl[:text])}</a:t></a:r></a:p></c:rich></c:tx>" if lbl[:text]
                if lbl[:num_fmt]
                  nf = lbl[:num_fmt]
                  nf_src = nf.is_a?(Hash) ? nf : { format_code: nf }
                  nf_linked = if nf_src.key?(:source_linked)
                                nf_src[:source_linked] ? 1 : 0
                              else
                                0
                              end
                  parts << %(<c:numFmt formatCode="#{xml_escape(nf_src[:format_code])}" sourceLinked="#{nf_linked}"/>)
                end
                tll_sp = +""
                tll_sp << %(<a:solidFill>#{color_xml(lbl[:fill_color])}</a:solidFill>) if lbl[:fill_color]
                tll_sp << "<a:noFill/>" if lbl[:no_fill]
                if lbl[:line_color] || lbl[:line_width] || lbl[:line_dash]
                  tll_lw = lbl[:line_width] ? %( w="#{(lbl[:line_width] * 12_700).to_i}") : ""
                  tll_lf = lbl[:line_color] ? %(<a:solidFill>#{color_xml(lbl[:line_color])}</a:solidFill>) : ""
                  tll_ld = lbl[:line_dash] ? %(<a:prstDash val="#{xml_escape(lbl[:line_dash])}"/>) : ""
                  tll_sp << "<a:ln#{tll_lw}>#{tll_lf}#{tll_ld}</a:ln>"
                end
                parts << "<c:spPr>#{tll_sp}</c:spPr>" unless tll_sp.empty?
                parts << build_axis_txpr(nil, lbl[:font]) if lbl[:font]
                parts << "</c:trendlineLbl>"
              end
              parts << "</c:trendline>"
            end
            err_bars_list = ser[:error_bars_list] || (ser[:error_bars] ? [ser[:error_bars]] : [])
            err_bars_list.each do |eb|
              parts << "<c:errBars>"
              parts << %(<c:errDir val="#{xml_escape(eb[:direction])}"/>) if eb[:direction]
              parts << %(<c:errBarType val="#{xml_escape(eb[:bar_type] || "both")}"/>)
              parts << %(<c:errValType val="#{xml_escape(eb[:val_type] || "fixedVal")}"/>)
              parts << %(<c:noEndCap val="#{eb[:no_end_cap] ? 1 : 0}"/>) unless eb[:no_end_cap].nil?
              parts << "<c:plus><c:numRef><c:f>#{xml_escape(eb[:plus])}</c:f></c:numRef></c:plus>" if eb[:plus]
              parts << "<c:minus><c:numRef><c:f>#{xml_escape(eb[:minus])}</c:f></c:numRef></c:minus>" if eb[:minus]
              parts << %(<c:val val="#{eb[:val]}"/>) if eb[:val]
              eb_sp_children = +""
              eb_sp_children << %(<a:solidFill>#{color_xml(eb[:fill_color])}</a:solidFill>) if eb[:fill_color]
              eb_sp_children << "<a:noFill/>" if eb[:no_fill]
              if eb[:line_color] || eb[:line_width] || eb[:line_dash]
                eb_ln_w = eb[:line_width] ? %( w="#{(eb[:line_width] * 12_700).to_i}") : ""
                eb_ln_f = eb[:line_color] ? %(<a:solidFill>#{color_xml(eb[:line_color])}</a:solidFill>) : ""
                eb_ln_d = eb[:line_dash] ? %(<a:prstDash val="#{xml_escape(eb[:line_dash])}"/>) : ""
                eb_sp_children << "<a:ln#{eb_ln_w}>#{eb_ln_f}#{eb_ln_d}</a:ln>"
              end
              parts << "<c:spPr>#{eb_sp_children}</c:spPr>" unless eb_sp_children.empty?
              parts << "</c:errBars>"
            end
            uses_xy = %w[scatterChart bubbleChart].include?(chart_type)
            cat_tag = uses_xy ? "xVal" : "cat"
            val_tag = uses_xy ? "yVal" : "val"
            if ser[:cat_ref]
              parts << if uses_xy || ser[:cat_ref_type] == :num
                         "<c:#{cat_tag}><c:numRef><c:f>#{xml_escape(ser[:cat_ref])}</c:f>#{num_cache_xml(ser[:cat_ref])}</c:numRef></c:#{cat_tag}>"
                       else
                         "<c:#{cat_tag}><c:strRef><c:f>#{xml_escape(ser[:cat_ref])}</c:f>#{str_cache_xml(ser[:cat_ref])}</c:strRef></c:#{cat_tag}>"
                       end
            end
            parts << "<c:#{val_tag}><c:numRef><c:f>#{xml_escape(ser[:val_ref])}</c:f>#{num_cache_xml(ser[:val_ref])}</c:numRef></c:#{val_tag}>" if ser[:val_ref]
            parts << "<c:bubbleSize><c:numRef><c:f>#{xml_escape(ser[:bubble_size_ref])}</c:f>#{num_cache_xml(ser[:bubble_size_ref])}</c:numRef></c:bubbleSize>" if ser[:bubble_size_ref]
            parts << %(<c:smooth val="#{ser[:smooth] ? 1 : 0}"/>) unless ser[:smooth].nil?
            parts << %(<c:shape val="#{xml_escape(ser[:shape])}"/>) if ser[:shape]
            parts << "</c:ser>"
          end

          if chart[:band_fmts]
            parts << "<c:bandFmts>"
            chart[:band_fmts].each do |bf|
              parts << "<c:bandFmt><c:idx val=\"#{bf[:idx]}\"/>"
              bf_sp = +""
              bf_sp << %(<a:solidFill>#{color_xml(bf[:fill_color])}</a:solidFill>) if bf[:fill_color]
              bf_sp << "<a:noFill/>" if bf[:no_fill]
              if bf[:line_color] || bf[:line_width] || bf[:line_dash]
                bf_lw = bf[:line_width] ? %( w="#{(bf[:line_width] * 12_700).to_i}") : ""
                bf_lf = bf[:line_color] ? %(<a:solidFill>#{color_xml(bf[:line_color])}</a:solidFill>) : ""
                bf_ld = bf[:line_dash] ? %(<a:prstDash val="#{xml_escape(bf[:line_dash])}"/>) : ""
                bf_sp << "<a:ln#{bf_lw}>#{bf_lf}#{bf_ld}</a:ln>"
              end
              parts << "<c:spPr>#{bf_sp}</c:spPr>" unless bf_sp.empty?
              parts << "</c:bandFmt>"
            end
            parts << "</c:bandFmts>"
          end

          parts << %(<c:gapWidth val="#{chart[:gap_width]}"/>) if chart[:gap_width]
          parts << %(<c:splitType val="#{xml_escape(chart[:split_type])}"/>) if chart[:split_type]
          parts << %(<c:splitPos val="#{chart[:split_pos]}"/>) if chart[:split_pos]
          if chart[:cust_split]&.any?
            parts << "<c:custSplit>"
            chart[:cust_split].each { |idx| parts << %(<c:secondPiePt val="#{idx}"/>) }
            parts << "</c:custSplit>"
          end
          parts << %(<c:secondPieSize val="#{chart[:second_pie_size]}"/>) if chart[:second_pie_size]
          parts << %(<c:gapDepth val="#{chart[:gap_depth]}"/>) if chart[:gap_depth]
          parts << %(<c:overlap val="#{chart[:overlap]}"/>) if chart[:overlap]
          parts << %(<c:shape val="#{chart[:bar_shape]}"/>) if chart[:bar_shape]
          b3d = chart[:bubble_3d]
          parts << %(<c:bubble3D val="#{b3d ? 1 : 0}"/>) unless b3d.nil?
          parts << %(<c:bubbleScale val="#{chart[:bubble_scale]}"/>) if chart[:bubble_scale]
          snb = chart[:show_neg_bubbles]
          parts << %(<c:showNegBubbles val="#{snb ? 1 : 0}"/>) unless snb.nil?
          parts << %(<c:sizeRepresents val="#{chart[:size_represents]}"/>) if chart[:size_represents]
          parts << %(<c:firstSliceAng val="#{chart[:first_slice_ang]}"/>) if chart[:first_slice_ang]
          parts << %(<c:holeSize val="#{chart[:hole_size]}"/>) if chart[:hole_size]
          if chart[:ser_lines]
            sl_spec = chart[:ser_lines]
            if sl_spec.is_a?(Hash)
              sl_sp = +""
              if sl_spec[:line_color] || sl_spec[:line_width] || sl_spec[:line_dash]
                sl_lw = sl_spec[:line_width] ? %( w="#{(sl_spec[:line_width] * 12_700).to_i}") : ""
                sl_lf = sl_spec[:line_color] ? %(<a:solidFill>#{color_xml(sl_spec[:line_color])}</a:solidFill>) : ""
                sl_ld = sl_spec[:line_dash] ? %(<a:prstDash val="#{xml_escape(sl_spec[:line_dash])}"/>) : ""
                sl_sp << "<a:ln#{sl_lw}>#{sl_lf}#{sl_ld}</a:ln>"
              end
              parts << (sl_sp.empty? ? "<c:serLines/>" : "<c:serLines><c:spPr>#{sl_sp}</c:spPr></c:serLines>")
            else
              parts << "<c:serLines/>"
            end
          end
          if chart[:drop_lines]
            dl_spec = chart[:drop_lines]
            if dl_spec.is_a?(Hash)
              dl_sp = +""
              if dl_spec[:line_color] || dl_spec[:line_width] || dl_spec[:line_dash]
                dl_lw = dl_spec[:line_width] ? %( w="#{(dl_spec[:line_width] * 12_700).to_i}") : ""
                dl_lf = dl_spec[:line_color] ? %(<a:solidFill>#{color_xml(dl_spec[:line_color])}</a:solidFill>) : ""
                dl_ld = dl_spec[:line_dash] ? %(<a:prstDash val="#{xml_escape(dl_spec[:line_dash])}"/>) : ""
                dl_sp << "<a:ln#{dl_lw}>#{dl_lf}#{dl_ld}</a:ln>"
              end
              parts << (dl_sp.empty? ? "<c:dropLines/>" : "<c:dropLines><c:spPr>#{dl_sp}</c:spPr></c:dropLines>")
            else
              parts << "<c:dropLines/>"
            end
          end
          if chart[:hi_low_lines]
            hl_spec = chart[:hi_low_lines]
            if hl_spec.is_a?(Hash)
              hl_sp = +""
              if hl_spec[:line_color] || hl_spec[:line_width] || hl_spec[:line_dash]
                hl_lw = hl_spec[:line_width] ? %( w="#{(hl_spec[:line_width] * 12_700).to_i}") : ""
                hl_lf = hl_spec[:line_color] ? %(<a:solidFill>#{color_xml(hl_spec[:line_color])}</a:solidFill>) : ""
                hl_ld = hl_spec[:line_dash] ? %(<a:prstDash val="#{xml_escape(hl_spec[:line_dash])}"/>) : ""
                hl_sp << "<a:ln#{hl_lw}>#{hl_lf}#{hl_ld}</a:ln>"
              end
              parts << (hl_sp.empty? ? "<c:hiLowLines/>" : "<c:hiLowLines><c:spPr>#{hl_sp}</c:spPr></c:hiLowLines>")
            else
              parts << "<c:hiLowLines/>"
            end
          end
          if chart[:up_down_bars]
            udb = chart[:up_down_bars]
            parts << "<c:upDownBars>"
            parts << %(<c:gapWidth val="#{udb[:gap_width]}"/>) if udb.is_a?(Hash) && udb[:gap_width]
            %i[up_bars down_bars].each do |bar_key|
              tag = bar_key == :up_bars ? "upBars" : "downBars"
              bar = udb.is_a?(Hash) ? udb[bar_key] : nil
              if bar
                bar_sp = +""
                bar_sp << %(<a:solidFill>#{color_xml(bar[:fill_color])}</a:solidFill>) if bar[:fill_color]
                bar_sp << "<a:noFill/>" if bar[:no_fill]
                if bar[:line_color] || bar[:line_width] || bar[:line_dash]
                  b_lw = bar[:line_width] ? %( w="#{(bar[:line_width] * 12_700).to_i}") : ""
                  b_lf = bar[:line_color] ? %(<a:solidFill>#{color_xml(bar[:line_color])}</a:solidFill>) : ""
                  b_ld = bar[:line_dash] ? %(<a:prstDash val="#{xml_escape(bar[:line_dash])}"/>) : ""
                  bar_sp << "<a:ln#{b_lw}>#{b_lf}#{b_ld}</a:ln>"
                end
                parts << (bar_sp.empty? ? "<c:#{tag}/>" : "<c:#{tag}><c:spPr>#{bar_sp}</c:spPr></c:#{tag}>")
              else
                parts << "<c:#{tag}/>"
              end
            end
            parts << "</c:upDownBars>"
          end
          mk = chart[:marker]
          parts << %(<c:marker val="#{mk ? 1 : 0}"/>) unless mk.nil?
          sm = chart[:smooth]
          parts << %(<c:smooth val="#{sm ? 1 : 0}"/>) unless sm.nil?
          three_axes = THREE_AXIS_CHARTS.include?(chart_type)
          unless no_axes
            parts << '<c:axId val="1"/><c:axId val="2"/>'
            parts << '<c:axId val="3"/>' if three_axes
          end
          parts << "</c:#{chart_type}>"

          unless no_axes
            cat_del = chart[:cat_axis_delete] ? 1 : 0
            cat_orient = chart[:cat_axis_orientation] || "minMax"
            cat_ax_tag = chart[:cat_axis_type] == :date ? "dateAx" : "catAx"
            parts << %(<c:#{cat_ax_tag}><c:axId val="1"/><c:scaling>)
            parts << %(<c:logBase val="#{chart[:cat_axis_log_base]}"/>) if chart[:cat_axis_log_base]
            parts << %(<c:orientation val="#{cat_orient}"/>)
            parts << %(<c:max val="#{chart[:cat_axis_scaling_max]}"/>) if chart[:cat_axis_scaling_max]
            parts << %(<c:min val="#{chart[:cat_axis_scaling_min]}"/>) if chart[:cat_axis_scaling_min]
            parts << %(</c:scaling><c:delete val="#{cat_del}"/><c:axPos val="#{chart[:cat_axis_pos] || "b"}"/>)
            parts << gridlines_xml("majorGridlines", chart[:cat_axis_major_gridlines])
            parts << gridlines_xml("minorGridlines", chart[:cat_axis_minor_gridlines])
            if chart[:cat_axis_title]
              cat_title_spec = merge_flat_title_styling(chart[:cat_axis_title], chart, :cat_axis_title)
              parts << build_chart_title_xml(cat_title_spec)
            end
            if (cnf = chart[:cat_axis_num_fmt])
              sl = cnf[:source_linked] ? 1 : 0
              parts << %(<c:numFmt formatCode="#{xml_escape(cnf[:format_code])}" sourceLinked="#{sl}"/>)
            end
            parts << %(<c:majorTickMark val="#{chart[:cat_axis_major_tick_mark]}"/>) if chart[:cat_axis_major_tick_mark]
            parts << %(<c:minorTickMark val="#{chart[:cat_axis_minor_tick_mark]}"/>) if chart[:cat_axis_minor_tick_mark]
            parts << %(<c:tickLblPos val="#{chart[:cat_axis_tick_lbl_pos]}"/>) if chart[:cat_axis_tick_lbl_pos]
            if chart[:cat_axis_line_color] || chart[:cat_axis_fill] || chart[:cat_axis_no_fill] || chart[:cat_axis_line_dash]
              parts << "<c:spPr>"
              parts << "<a:noFill/>" if chart[:cat_axis_no_fill]
              parts << %(<a:solidFill>#{color_xml(chart[:cat_axis_fill])}</a:solidFill>) if chart[:cat_axis_fill]
              if chart[:cat_axis_line_color] || chart[:cat_axis_line_dash]
                w_attr = chart[:cat_axis_line_width] ? %( w="#{chart[:cat_axis_line_width]}") : ""
                ca_ln_f = chart[:cat_axis_line_color] ? %(<a:solidFill>#{srgb_clr_xml(chart[:cat_axis_line_color])}</a:solidFill>) : ""
                ca_ln_d = chart[:cat_axis_line_dash] ? %(<a:prstDash val="#{xml_escape(chart[:cat_axis_line_dash])}"/>) : ""
                parts << "<a:ln#{w_attr}>#{ca_ln_f}#{ca_ln_d}</a:ln>"
              end
              parts << "</c:spPr>"
            end
            parts << build_axis_txpr(chart[:cat_axis_label_rotation], chart[:cat_axis_font])
            parts << '<c:crossAx val="2"/>'
            parts << %(<c:crosses val="#{chart[:cat_axis_crosses]}"/>) if chart[:cat_axis_crosses]
            parts << %(<c:crossesAt val="#{chart[:cat_axis_crosses_at]}"/>) if !chart[:cat_axis_crosses] && chart[:cat_axis_crosses_at]
            ca_auto = chart[:cat_axis_auto]
            parts << %(<c:auto val="#{ca_auto ? 1 : 0}"/>) unless ca_auto.nil?
            parts << %(<c:lblAlgn val="#{xml_escape(chart[:cat_axis_lbl_algn])}"/>) if chart[:cat_axis_lbl_algn]
            parts << %(<c:lblOffset val="#{chart[:cat_axis_lbl_offset]}"/>) if chart[:cat_axis_lbl_offset]
            parts << %(<c:tickLblSkip val="#{chart[:cat_axis_tick_lbl_skip]}"/>) if chart[:cat_axis_tick_lbl_skip]
            parts << %(<c:tickMarkSkip val="#{chart[:cat_axis_tick_mark_skip]}"/>) if chart[:cat_axis_tick_mark_skip]
            nml = chart[:cat_axis_no_multi_lvl_lbl]
            parts << %(<c:noMultiLvlLbl val="#{nml ? 1 : 0}"/>) unless nml.nil?
            if cat_ax_tag == "dateAx"
              parts << %(<c:baseTimeUnit val="#{chart[:cat_axis_base_time_unit]}"/>) if chart[:cat_axis_base_time_unit]
              parts << %(<c:majorUnit val="#{chart[:cat_axis_major_unit]}"/>) if chart[:cat_axis_major_unit]
              parts << %(<c:majorTimeUnit val="#{chart[:cat_axis_major_time_unit]}"/>) if chart[:cat_axis_major_time_unit]
              parts << %(<c:minorUnit val="#{chart[:cat_axis_minor_unit]}"/>) if chart[:cat_axis_minor_unit]
              parts << %(<c:minorTimeUnit val="#{chart[:cat_axis_minor_time_unit]}"/>) if chart[:cat_axis_minor_time_unit]
            end
            parts << "</c:#{cat_ax_tag}>"
            val_del = chart[:val_axis_delete] ? 1 : 0
            val_orient = chart[:val_axis_orientation] || "minMax"
            parts << %(<c:valAx><c:axId val="2"/><c:scaling>)
            parts << %(<c:logBase val="#{chart[:val_axis_log_base]}"/>) if chart[:val_axis_log_base]
            parts << %(<c:orientation val="#{val_orient}"/>)
            parts << %(<c:max val="#{chart[:val_axis_scaling_max]}"/>) if chart[:val_axis_scaling_max]
            parts << %(<c:min val="#{chart[:val_axis_scaling_min]}"/>) if chart[:val_axis_scaling_min]
            parts << %(</c:scaling><c:delete val="#{val_del}"/><c:axPos val="#{chart[:val_axis_pos] || "l"}"/>)
            parts << gridlines_xml("majorGridlines", chart[:val_axis_major_gridlines])
            parts << gridlines_xml("minorGridlines", chart[:val_axis_minor_gridlines])
            if chart[:val_axis_title]
              val_title_spec = merge_flat_title_styling(chart[:val_axis_title], chart, :val_axis_title)
              parts << build_chart_title_xml(val_title_spec)
            end
            if (vnf = chart[:val_axis_num_fmt])
              sl = vnf[:source_linked] ? 1 : 0
              parts << %(<c:numFmt formatCode="#{xml_escape(vnf[:format_code])}" sourceLinked="#{sl}"/>)
            end
            parts << %(<c:majorTickMark val="#{chart[:val_axis_major_tick_mark]}"/>) if chart[:val_axis_major_tick_mark]
            parts << %(<c:minorTickMark val="#{chart[:val_axis_minor_tick_mark]}"/>) if chart[:val_axis_minor_tick_mark]
            parts << %(<c:tickLblPos val="#{chart[:val_axis_tick_lbl_pos]}"/>) if chart[:val_axis_tick_lbl_pos]
            if chart[:val_axis_line_color] || chart[:val_axis_fill] || chart[:val_axis_no_fill] || chart[:val_axis_line_dash]
              parts << "<c:spPr>"
              parts << "<a:noFill/>" if chart[:val_axis_no_fill]
              parts << %(<a:solidFill>#{color_xml(chart[:val_axis_fill])}</a:solidFill>) if chart[:val_axis_fill]
              if chart[:val_axis_line_color] || chart[:val_axis_line_dash]
                w_attr = chart[:val_axis_line_width] ? %( w="#{chart[:val_axis_line_width]}") : ""
                va_ln_f = chart[:val_axis_line_color] ? %(<a:solidFill>#{srgb_clr_xml(chart[:val_axis_line_color])}</a:solidFill>) : ""
                va_ln_d = chart[:val_axis_line_dash] ? %(<a:prstDash val="#{xml_escape(chart[:val_axis_line_dash])}"/>) : ""
                parts << "<a:ln#{w_attr}>#{va_ln_f}#{va_ln_d}</a:ln>"
              end
              parts << "</c:spPr>"
            end
            parts << build_axis_txpr(chart[:val_axis_label_rotation], chart[:val_axis_font])
            parts << '<c:crossAx val="1"/>'
            parts << %(<c:crosses val="#{chart[:val_axis_crosses]}"/>) if chart[:val_axis_crosses]
            parts << %(<c:crossesAt val="#{chart[:val_axis_crosses_at]}"/>) if !chart[:val_axis_crosses] && chart[:val_axis_crosses_at]
            parts << %(<c:crossBetween val="#{chart[:val_axis_cross_between]}"/>) if chart[:val_axis_cross_between]
            parts << %(<c:majorUnit val="#{chart[:val_axis_major_unit]}"/>) if chart[:val_axis_major_unit]
            parts << %(<c:minorUnit val="#{chart[:val_axis_minor_unit]}"/>) if chart[:val_axis_minor_unit]
            if chart[:val_axis_disp_units]
              du = chart[:val_axis_disp_units]
              if du.is_a?(Hash)
                parts << "<c:dispUnits>"
                parts << %(<c:builtInUnit val="#{xml_escape(du[:built_in_unit])}"/>) if du[:built_in_unit]
                parts << %(<c:custUnit val="#{du[:cust_unit]}"/>) if du[:cust_unit]
                if du[:label]
                  dul = du[:label]
                  parts << "<c:dispUnitsLbl>"
                  if dul[:num_fmt]
                    nf = dul[:num_fmt]
                    nf_src = nf.is_a?(Hash) ? nf : { format_code: nf }
                    nf_linked = if nf_src.key?(:source_linked)
                                  nf_src[:source_linked] ? 1 : 0
                                else
                                  0
                                end
                    parts << %(<c:numFmt formatCode="#{xml_escape(nf_src[:format_code])}" sourceLinked="#{nf_linked}"/>)
                  end
                  dul_sp = +""
                  dul_sp << %(<a:solidFill>#{color_xml(dul[:fill_color])}</a:solidFill>) if dul[:fill_color]
                  dul_sp << "<a:noFill/>" if dul[:no_fill]
                  if dul[:line_color] || dul[:line_width] || dul[:line_dash]
                    dul_lw = dul[:line_width] ? %( w="#{(dul[:line_width] * 12_700).to_i}") : ""
                    dul_lf = dul[:line_color] ? %(<a:solidFill>#{color_xml(dul[:line_color])}</a:solidFill>) : ""
                    dul_ld = dul[:line_dash] ? %(<a:prstDash val="#{xml_escape(dul[:line_dash])}"/>) : ""
                    dul_sp << "<a:ln#{dul_lw}>#{dul_lf}#{dul_ld}</a:ln>"
                  end
                  parts << "<c:spPr>#{dul_sp}</c:spPr>" unless dul_sp.empty?
                  parts << build_axis_txpr(nil, dul[:font]) if dul[:font]
                  parts << "</c:dispUnitsLbl>"
                end
                parts << "</c:dispUnits>"
              else
                parts << %(<c:dispUnits><c:builtInUnit val="#{xml_escape(du.to_s)}"/></c:dispUnits>)
              end
            end
            parts << "</c:valAx>"

            if three_axes
              parts << '<c:serAx><c:axId val="3"/><c:scaling><c:orientation val="minMax"/></c:scaling>'
              parts << '<c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:serAx>'
            end
          end

          if chart[:data_table]
            dt = chart[:data_table]
            parts << "<c:dTable>"
            parts << %(<c:showHorzBorder val="#{dt[:show_horz_border] ? 1 : 0}"/>) unless dt[:show_horz_border].nil?
            parts << %(<c:showVertBorder val="#{dt[:show_vert_border] ? 1 : 0}"/>) unless dt[:show_vert_border].nil?
            parts << %(<c:showOutline val="#{dt[:show_outline] ? 1 : 0}"/>) unless dt[:show_outline].nil?
            parts << %(<c:showKeys val="#{dt[:show_keys] ? 1 : 0}"/>) unless dt[:show_keys].nil?
            dt_sp_children = +""
            dt_sp_children << %(<a:solidFill>#{color_xml(dt[:fill_color])}</a:solidFill>) if dt[:fill_color]
            dt_sp_children << "<a:noFill/>" if dt[:no_fill]
            if dt[:line_color] || dt[:line_width] || dt[:line_dash]
              dt_ln_w = dt[:line_width] ? %( w="#{(dt[:line_width] * 12_700).to_i}") : ""
              dt_ln_f = dt[:line_color] ? %(<a:solidFill>#{color_xml(dt[:line_color])}</a:solidFill>) : ""
              dt_ln_d = dt[:line_dash] ? %(<a:prstDash val="#{xml_escape(dt[:line_dash])}"/>) : ""
              dt_sp_children << "<a:ln#{dt_ln_w}>#{dt_ln_f}#{dt_ln_d}</a:ln>"
            end
            parts << "<c:spPr>#{dt_sp_children}</c:spPr>" unless dt_sp_children.empty?
            parts << build_axis_txpr(nil, dt[:font]) if dt[:font]
            parts << "</c:dTable>"
          end

          if chart[:plot_area_fill] || chart[:plot_area_no_fill] || chart[:plot_area_line_color] || chart[:plot_area_line_dash]
            pa_sp = +""
            pa_sp << %(<a:solidFill>#{color_xml(chart[:plot_area_fill])}</a:solidFill>) if chart[:plot_area_fill]
            pa_sp << "<a:noFill/>" if chart[:plot_area_no_fill]
            if chart[:plot_area_line_color] || chart[:plot_area_line_dash]
              pa_ln_w = chart[:plot_area_line_width] ? %( w="#{chart[:plot_area_line_width].to_i}") : ""
              pa_ln_f = chart[:plot_area_line_color] ? %(<a:solidFill>#{color_xml(chart[:plot_area_line_color])}</a:solidFill>) : ""
              pa_ln_d = chart[:plot_area_line_dash] ? %(<a:prstDash val="#{xml_escape(chart[:plot_area_line_dash])}"/>) : ""
              pa_sp << "<a:ln#{pa_ln_w}>#{pa_ln_f}#{pa_ln_d}</a:ln>"
            end
            parts << "<c:spPr>#{pa_sp}</c:spPr>"
          end

          parts << "</c:plotArea>"
          legend_pos = chart.dig(:legend, :position) || "r"
          legend_overlay = chart.dig(:legend, :overlay)
          parts << %(<c:legend><c:legendPos val="#{legend_pos}"/>)
          legend_entries = chart.dig(:legend, :entries)
          legend_entries&.each do |entry|
            parts << %(<c:legendEntry><c:idx val="#{entry[:idx]}"/>)
            if !entry[:delete].nil?
              parts << %(<c:delete val="#{entry[:delete] ? 1 : 0}"/>)
            elsif entry[:font]
              parts << build_axis_txpr(nil, entry[:font])
            end
            parts << "</c:legendEntry>"
          end
          legend_layout = chart.dig(:legend, :layout)
          if legend_layout.is_a?(Hash)
            ml_parts = +""
            ml_parts << %(<c:layoutTarget val="#{xml_escape(legend_layout[:target])}"/>) if legend_layout[:target]
            ml_parts << %(<c:xMode val="#{xml_escape(legend_layout[:x_mode])}"/>) if legend_layout[:x_mode]
            ml_parts << %(<c:yMode val="#{xml_escape(legend_layout[:y_mode])}"/>) if legend_layout[:y_mode]
            ml_parts << %(<c:wMode val="#{xml_escape(legend_layout[:w_mode])}"/>) if legend_layout[:w_mode]
            ml_parts << %(<c:hMode val="#{xml_escape(legend_layout[:h_mode])}"/>) if legend_layout[:h_mode]
            ml_parts << %(<c:x val="#{legend_layout[:x]}"/>) if legend_layout[:x]
            ml_parts << %(<c:y val="#{legend_layout[:y]}"/>) if legend_layout[:y]
            ml_parts << %(<c:w val="#{legend_layout[:w]}"/>) if legend_layout[:w]
            ml_parts << %(<c:h val="#{legend_layout[:h]}"/>) if legend_layout[:h]
            parts << "<c:layout><c:manualLayout>#{ml_parts}</c:manualLayout></c:layout>" unless ml_parts.empty?
          end
          parts << %(<c:overlay val="#{legend_overlay ? 1 : 0}"/>) unless legend_overlay.nil?
          leg_sp_children = +""
          leg_fill = chart.dig(:legend, :fill_color)
          leg_sp_children << %(<a:solidFill>#{color_xml(leg_fill)}</a:solidFill>) if leg_fill
          leg_no_fill = chart.dig(:legend, :no_fill)
          leg_sp_children << "<a:noFill/>" if leg_no_fill
          leg_lc = chart.dig(:legend, :line_color)
          leg_lw = chart.dig(:legend, :line_width)
          leg_ld = chart.dig(:legend, :line_dash)
          if leg_lc || leg_lw || leg_ld
            lw_attr = leg_lw ? %( w="#{(leg_lw * 12_700).to_i}") : ""
            lf = leg_lc ? %(<a:solidFill>#{color_xml(leg_lc)}</a:solidFill>) : ""
            ld = leg_ld ? %(<a:prstDash val="#{xml_escape(leg_ld)}"/>) : ""
            leg_sp_children << "<a:ln#{lw_attr}>#{lf}#{ld}</a:ln>"
          end
          parts << "<c:spPr>#{leg_sp_children}</c:spPr>" unless leg_sp_children.empty?
          if (lfont = chart.dig(:legend, :font))
            parts << build_axis_txpr(nil, lfont)
          end
          parts << %(</c:legend>)
          pvo = chart[:plot_vis_only]
          parts << %(<c:plotVisOnly val="#{pvo ? 1 : 0}"/>) unless pvo.nil?
          parts << %(<c:dispBlanksAs val="#{chart[:disp_blanks_as]}"/>) if chart[:disp_blanks_as]
          sdom = chart[:show_d_lbls_over_max]
          parts << %(<c:showDLblsOverMax val="#{sdom ? 1 : 0}"/>) unless sdom.nil?
          parts << "</c:chart>"
          if chart[:chart_fill] || chart[:chart_no_fill] || chart[:chart_line_color] || chart[:chart_line_dash]
            cs_sp = +""
            cs_sp << %(<a:solidFill>#{color_xml(chart[:chart_fill])}</a:solidFill>) if chart[:chart_fill]
            cs_sp << "<a:noFill/>" if chart[:chart_no_fill]
            if chart[:chart_line_color] || chart[:chart_line_dash]
              cs_lw = chart[:chart_line_width] ? %( w="#{(chart[:chart_line_width] * 12_700).to_i}") : ""
              cs_ln_f = chart[:chart_line_color] ? %(<a:solidFill>#{color_xml(chart[:chart_line_color])}</a:solidFill>) : ""
              cs_ln_d = chart[:chart_line_dash] ? %(<a:prstDash val="#{xml_escape(chart[:chart_line_dash])}"/>) : ""
              cs_sp << "<a:ln#{cs_lw}>#{cs_ln_f}#{cs_ln_d}</a:ln>"
            end
            parts << "<c:spPr>#{cs_sp}</c:spPr>"
          end
          parts << build_axis_txpr(nil, chart[:chart_font]) if chart[:chart_font]
          if (ps = chart[:print_settings])
            parts << "<c:printSettings>"
            if (hf = ps[:header_footer])
              parts << "<c:headerFooter>"
              parts << "<c:oddHeader>#{xml_escape(hf[:odd_header])}</c:oddHeader>" if hf[:odd_header]
              parts << "<c:oddFooter>#{xml_escape(hf[:odd_footer])}</c:oddFooter>" if hf[:odd_footer]
              parts << "<c:evenHeader>#{xml_escape(hf[:even_header])}</c:evenHeader>" if hf[:even_header]
              parts << "<c:evenFooter>#{xml_escape(hf[:even_footer])}</c:evenFooter>" if hf[:even_footer]
              parts << "<c:firstHeader>#{xml_escape(hf[:first_header])}</c:firstHeader>" if hf[:first_header]
              parts << "<c:firstFooter>#{xml_escape(hf[:first_footer])}</c:firstFooter>" if hf[:first_footer]
              parts << "</c:headerFooter>"
            end
            if (pm = ps[:page_margins])
              pm_attrs = %w[b l r t header footer].filter_map { |a| pm[a.to_sym] ? %( #{a}="#{pm[a.to_sym]}") : nil }.join
              parts << "<c:pageMargins#{pm_attrs}/>"
            end
            if (psu = ps[:page_setup])
              psu_parts = +""
              psu_parts << %( paperSize="#{psu[:paper_size]}") if psu[:paper_size]
              psu_parts << %( firstPageNumber="#{psu[:first_page_number]}") if psu[:first_page_number]
              psu_parts << %( orientation="#{xml_escape(psu[:orientation])}") if psu[:orientation]
              psu_parts << %( horizontalDpi="#{psu[:horizontal_dpi]}") if psu[:horizontal_dpi]
              psu_parts << %( verticalDpi="#{psu[:vertical_dpi]}") if psu[:vertical_dpi]
              psu_parts << %( copies="#{psu[:copies]}") if psu[:copies]
              parts << "<c:pageSetup#{psu_parts}/>"
            end
            parts << "</c:printSettings>"
          end
          parts << "</c:chartSpace>"
          parts.join
        end

        # : (untyped sheet_comments) -> untyped
        def generate_comments_xml(sheet_comments)
          authors = sheet_comments.map { |c| c[:author] }.uniq
          parts = [
            XML_HEADER,
            %(<comments xmlns="#{SSML_NS}">),
            "<authors>"
          ]
          authors.each { |a| parts << "<author>#{xml_escape(a)}</author>" }
          parts << "</authors><commentList>"
          sheet_comments.each do |c|
            aid = authors.index(c[:author]) || 0
            text_xml = if c[:text].is_a?(Xlsxrb::Elements::RichText)
                         rich_text_xml(c[:text])
                       else
                         "<r><t>#{xml_escape(c[:text])}</t></r>"
                       end
            comment_attrs = %(ref="#{c[:ref]}" authorId="#{aid}")
            comment_attrs << %( guid="#{c[:guid]}") if c[:guid]
            comment_attrs << %( shapeId="#{c[:shape_id]}") if c[:shape_id]
            parts << "<comment #{comment_attrs}><text>#{text_xml}</text></comment>"
          end
          parts << "</commentList></comments>"
          parts.join
        end

        # : (untyped sheet_comments) -> untyped
        def generate_vml_drawing_xml(sheet_comments)
          parts = [
            '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
            '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
            '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
            '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/>',
            "</v:shapetype>"
          ]
          sheet_comments.each_with_index do |c, idx|
            col, row = cell_to_col_row(c[:ref])
            shape_id = 1025 + idx
            parts << %(<v:shape id="_x0000_s#{shape_id}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:#{idx + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">)
            parts << '<v:fill color2="#ffffe1"/>'
            parts << '<v:shadow on="t" color="black" obscured="t"/>'
            parts << '<v:path o:connecttype="none"/>'
            parts << '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>'
            parts << '<x:ClientData ObjectType="Note">'
            parts << "<x:MoveWithCells/><x:SizeWithCells/>"
            parts << "<x:Anchor>#{col + 1}, 15, #{row}, 10, #{col + 3}, 15, #{row + 4}, 4</x:Anchor>"
            parts << "<x:AutoFill>False</x:AutoFill>"
            parts << "<x:Row>#{row}</x:Row>"
            parts << "<x:Column>#{col}</x:Column>"
            parts << "</x:ClientData></v:shape>"
          end
          parts << "</xml>"
          parts.join
        end

        # : (untyped cell_ref) -> (::Array[0] | ::Array[untyped])
        def cell_to_col_row(cell_ref)
          m = cell_ref.match(/\A([A-Z]+)(\d+)\z/)
          return [0, 0] unless m

          col = column_letter_to_index(m[1]) - 1
          row = m[2].to_i - 1
          [col, row]
        end
      end
    end
  end
end
