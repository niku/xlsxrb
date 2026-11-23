# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    class Writer
      # Mixin containing DrawingML, Chart, Shape, Comment, and VML XML generation logic.
      module DrawingXml
        # : (untyped drawing_parts) -> ::String
        def generate_drawing_xml(drawing_parts)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("xdr:wsDr", { "xmlns:xdr": XDR_NS, "xmlns:a": A_NS, "xmlns:r": DOC_REL_NS }) do
              drawing_parts.each do |dp|
                case dp[:kind]
                when :pic
                  write_pic_anchor_xml(xml, dp)
                when :chart
                  write_chart_anchor_xml(xml, dp)
                when :sp
                  write_shape_anchor_xml(xml, dp)
                end
              end
            end
          end
        end

        # : (untyped rels_data) -> ::String
        def generate_drawing_rels(rels_data)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("Relationships", { xmlns: REL_NS }) do
              rels_data.each_with_index do |rel, i|
                rel_type = case rel[:type]
                           when :image then "#{DOC_REL_NS}/image"
                           when :chart then "#{DOC_REL_NS}/chart"
                           end
                xml.empty_tag("Relationship", { Id: "rId#{i + 1}", Type: rel_type || "", Target: rel[:target] || "" })
              end
            end
          end
        end

        # : (untyped chart) -> ::String
        def generate_chart_xml(chart)
          # Charts read from files carry the element name ("lineChart") as their type.
          chart_type = CHART_TYPE_MAP[chart[:type]] || CHART_TYPE_MAP.values.find { |element| element == chart[:type].to_s } || "barChart"
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("c:chartSpace", { "xmlns:c": C_NS, "xmlns:a": A_NS, "xmlns:r": DOC_REL_NS }) do
              xml.empty_tag("c:roundedCorners", { val: chart[:rounded_corners] ? 1 : 0 }) unless chart[:rounded_corners].nil?
              xml.empty_tag("c:style", { val: chart[:style] }) if chart[:style]
              if (prot = chart[:protection])
                xml.tag("c:protection") do
                  xml.empty_tag("c:chartObject", { val: prot[:chart_object] ? 1 : 0 }) unless prot[:chart_object].nil?
                  xml.empty_tag("c:data", { val: prot[:data] ? 1 : 0 }) unless prot[:data].nil?
                  xml.empty_tag("c:formatting", { val: prot[:formatting] ? 1 : 0 }) unless prot[:formatting].nil?
                  xml.empty_tag("c:selection", { val: prot[:selection] ? 1 : 0 }) unless prot[:selection].nil?
                  xml.empty_tag("c:userInterface", { val: prot[:user_interface] ? 1 : 0 }) unless prot[:user_interface].nil?
                end
              end

              xml.tag("c:chart") do
                write_chart_top_title_xml(xml, chart)
                xml.empty_tag("c:autoTitleDeleted", { val: chart[:auto_title_deleted] ? 1 : 0 }) unless chart[:auto_title_deleted].nil?
                write_view_3d_xml(xml, chart[:view_3d]) if chart[:view_3d]
                write_chart_walls_xml(xml, chart)

                write_plot_area_xml(xml, chart, chart_type)

                write_chart_legend_xml(xml, chart[:legend])
                xml.empty_tag("c:plotVisOnly", { val: chart[:plot_vis_only] ? 1 : 0 }) unless chart[:plot_vis_only].nil?
                xml.empty_tag("c:dispBlanksAs", { val: chart[:disp_blanks_as] }) if chart[:disp_blanks_as]
                xml.empty_tag("c:showDLblsOverMax", { val: chart[:show_d_lbls_over_max] ? 1 : 0 }) unless chart[:show_d_lbls_over_max].nil?
              end

              if chart[:chart_fill] || chart[:chart_no_fill] || chart[:chart_line_color] || chart[:chart_line_dash]
                write_simple_sp_pr(
                  xml,
                  fill_color: chart[:chart_fill],
                  no_fill: chart[:chart_no_fill],
                  line_color: chart[:chart_line_color],
                  line_width: chart[:chart_line_width],
                  line_dash: chart[:chart_line_dash]
                )
              end
              write_axis_txpr(xml, nil, chart[:chart_font]) if chart[:chart_font]
              write_chart_print_settings_xml(xml, chart[:print_settings]) if chart[:print_settings]
            end
          end
        end

        # : (untyped sheet_comments) -> ::String
        def generate_comments_xml(sheet_comments)
          authors = sheet_comments.map { |c| c[:author] }.uniq
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("comments", { xmlns: SSML_NS }) do
              xml.tag("authors") do
                authors.each { |a| xml.text_tag("author", a || "") }
              end
              xml.tag("commentList") do
                sheet_comments.each do |c|
                  aid = authors.index(c[:author]) || 0
                  xml.tag("comment", { ref: c[:ref] || "", authorId: aid, guid: c[:guid], shapeId: c[:shape_id] }) do
                    xml.tag("text") do
                      if c[:text].is_a?(Xlsxrb::Elements::RichText)
                        write_rich_text_xml(xml, c[:text])
                      else
                        xml.tag("r") { xml.text_tag("t", c[:text] || "") }
                      end
                    end
                  end
                end
              end
            end
          end
        end

        # : (untyped sheet_comments) -> ::String
        def generate_vml_drawing_xml(sheet_comments)
          XmlBuilder.build do |xml|
            xml.tag(
              "xml",
              {
                "xmlns:v": "urn:schemas-microsoft-com:vml",
                "xmlns:o": "urn:schemas-microsoft-com:office:office",
                "xmlns:x": "urn:schemas-microsoft-com:office:excel"
              }
            ) do
              xml.tag("o:shapelayout", { "v:ext": "edit" }) do
                xml.empty_tag("o:idmap", { "v:ext": "edit", data: 1 })
              end
              xml.tag(
                "v:shapetype",
                {
                  id: "_x0000_t202",
                  coordsize: "21600,21600",
                  "o:spt": 202,
                  path: "m,l,21600r21600,l21600,xe"
                }
              ) do
                xml.empty_tag("v:stroke", { joinstyle: "miter" })
                xml.empty_tag("v:path", { gradientshapeok: "t", "o:connecttype": "rect" })
              end

              sheet_comments.each_with_index do |c, idx|
                col, row = cell_to_col_row(c[:ref])
                shape_id = 1025 + idx
                style = "position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:#{idx + 1};visibility:#{c[:visible] ? "visible" : "hidden"}"
                xml.tag(
                  "v:shape",
                  {
                    id: "_x0000_s#{shape_id}",
                    type: "#_x0000_t202",
                    style: style,
                    fillcolor: "#ffffe1",
                    "o:insetmode": "auto"
                  }
                ) do
                  xml.empty_tag("v:fill", { color2: "#ffffe1" })
                  xml.empty_tag("v:shadow", { on: "t", color: "black", obscured: "t" })
                  xml.empty_tag("v:path", { "o:connecttype": "none" })
                  xml.tag("v:textbox", { style: "mso-direction-alt:auto" }) do
                    xml.open_tag("div", { style: "text-align:left" })
                    xml.close_tag("div")
                  end
                  xml.tag("x:ClientData", { ObjectType: "Note" }) do
                    xml.empty_tag("x:MoveWithCells")
                    xml.empty_tag("x:SizeWithCells")
                    anchor = c[:anchor] || Xlsxrb::Elements::Comment.default_anchor(col, row)
                    xml.text_tag("x:Anchor", anchor.join(", "))
                    xml.text_tag("x:AutoFill", "False")
                    xml.text_tag("x:Row", row)
                    xml.text_tag("x:Column", col)
                    xml.empty_tag("x:Visible") if c[:visible]
                  end
                end
              end
            end
          end
        end

        # : (untyped cell_ref) -> (::Array[0] | ::Array[untyped])
        def cell_to_col_row(cell_ref)
          m = cell_ref.match(/\A([A-Z]+)(\d+)\z/)
          return [0, 0] unless m

          col = column_letter_to_index(m[1]) - 1
          row = m[2].to_i - 1
          [col, row]
        end

        THREADED_COMMENTS_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

        # Formats or generates a GUID in standard uppercase braced format {UUID}.
        # : (?String? val) -> String
        def format_guid(val = nil)
          return "{#{SecureRandom.uuid.upcase}}" if val.nil? || val.to_s.empty?

          str = val.to_s.strip
          str = "{#{str}}" unless str.start_with?("{") && str.end_with?("}")
          str.upcase
        end

        # Generates XML for xl/threadedComments/threadedComment{N}.xml.
        # : (Array[Hash[Symbol, untyped]] sheet_threaded_comments) -> String
        def generate_threaded_comments_xml(sheet_threaded_comments)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("ThreadedComments", { xmlns: THREADED_COMMENTS_NS }) do
              sheet_threaded_comments.each do |c|
                ref = c[:ref] || c[:cell]
                dt = c[:timestamp] || c[:dt]
                dt_str = case dt
                         when Time, DateTime
                           dt.to_time.utc.strftime("%Y-%m-%dT%H:%M:%SZ")
                         when String
                           dt
                         else
                           Time.now.utc.strftime("%Y-%m-%dT%H:%M:%SZ")
                         end
                person_id = c[:person_id] || c[:author_id] || "{00000000-0000-0000-0000-000000000000}"
                comment_id = c[:id] || format_guid(nil)
                attrs = {
                  ref: ref || "",
                  dT: dt_str,
                  personId: person_id,
                  id: comment_id,
                  parentId: c[:parent_id],
                  done: ([true, "true", "1"].include?(c[:done]) ? "true" : nil)
                }
                xml.tag("threadedComment", attrs) do
                  xml.text_tag("text", c[:text].to_s)
                end
              end
            end
          end
        end

        # Generates XML for xl/persons/person.xml.
        # : (Array[Hash[Symbol, untyped]] persons) -> String
        def generate_persons_xml(persons)
          XmlBuilder.build(declaration: :plain) do |xml|
            xml.tag("personList", { xmlns: THREADED_COMMENTS_NS }) do
              persons.each do |p|
                p_id = p[:id] || format_guid(nil)
                name = p[:display_name] || p[:name] || "Author"
                user_id = p[:user_id] || "AD:#{name}"
                provider_id = p[:provider_id] || "AD"
                xml.empty_tag("person", { displayName: name, id: p_id, userId: user_id, providerId: provider_id })
              end
            end
          end
        end

        private

        # : (XmlBuilder xml, untyped part) -> void
        def write_pic_anchor_xml(xml, part)
          img = part[:img]
          rid = "rId#{part[:rid_index]}"
          ea = img[:edit_as] || "oneCell"
          xml.tag("xdr:twoCellAnchor", { editAs: ea, fPublished: (1 if img[:published]) }) do
            write_anchor_xml(xml, "from", img[:from_column] || img[:from_col], img[:from_row], col_off: img[:from_col_off] || 0, row_off: img[:from_row_off] || 0)
            write_anchor_xml(xml, "to", img[:to_column] || img[:to_col], img[:to_row], col_off: img[:to_col_off] || 0, row_off: img[:to_row_off] || 0)
            xml.tag("xdr:pic", { macro: img[:macro] }) do
              xml.tag("xdr:nvPicPr") do
                xml.empty_tag(
                  "xdr:cNvPr",
                  {
                    id: part[:rid_index] + 1,
                    name: img[:name] || "",
                    descr: img[:description],
                    title: img[:title],
                    hidden: (1 if img[:hidden])
                  }
                )
                xml.tag("xdr:cNvPicPr") do
                  xml.empty_tag(
                    "a:picLocks",
                    {
                      noChangeAspect: (1 if img[:no_change_aspect]),
                      noCrop: (1 if img[:no_crop])
                    }
                  )
                end
              end
              xml.tag("xdr:blipFill") do
                if img[:alpha_mod_fix]
                  xml.tag("a:blip", { "r:embed": rid }) do
                    xml.empty_tag("a:alphaModFix", { amt: img[:alpha_mod_fix] })
                  end
                else
                  xml.empty_tag("a:blip", { "r:embed": rid })
                end
                if (sr = img[:src_rect])
                  xml.empty_tag("a:srcRect", { t: sr[:top], b: sr[:bottom], l: sr[:left], r: sr[:right] })
                end
                xml.tag("a:stretch") { xml.empty_tag("a:fillRect") }
              end
              xml.tag("xdr:spPr") do
                if img[:rotation]
                  xml.tag("a:xfrm", { rot: img[:rotation].to_i }) do
                    xml.empty_tag("a:off", { x: 0, y: 0 })
                    xml.empty_tag("a:ext", { cx: 0, cy: 0 })
                  end
                end
                xml.tag("a:prstGeom", { prst: "rect" }) { xml.empty_tag("a:avLst") }
                if img[:line_color]
                  xml.tag("a:ln", { w: img[:line_width]&.to_i }) do
                    xml.tag("a:solidFill") { write_drawing_color_xml(xml, img[:line_color]) }
                  end
                end
              end
            end
            write_client_data_xml(xml, img)
          end
        end

        # : (XmlBuilder xml, untyped part) -> void
        def write_chart_anchor_xml(xml, part)
          chart = part[:chart]
          rid = "rId#{part[:rid_index]}"
          xml.tag("xdr:twoCellAnchor", { editAs: chart[:edit_as], fPublished: (1 if chart[:published]) }) do
            write_anchor_xml(xml, "from", chart[:from_column] || chart[:from_col], chart[:from_row], col_off: chart[:from_col_off] || 0, row_off: chart[:from_row_off] || 0)
            write_anchor_xml(xml, "to", chart[:to_column] || chart[:to_col], chart[:to_row], col_off: chart[:to_col_off] || 0, row_off: chart[:to_row_off] || 0)
            xml.tag("xdr:graphicFrame", { macro: chart[:frame_macro] || "" }) do
              chart_frame_name = chart[:name] || chart[:title] || "Chart"
              xml.tag("xdr:nvGraphicFramePr") do
                xml.empty_tag(
                  "xdr:cNvPr",
                  {
                    id: part[:rid_index] + 1,
                    name: chart_frame_name,
                    descr: chart[:description],
                    title: chart[:frame_title],
                    hidden: (1 if chart[:frame_hidden])
                  }
                )
                if chart[:frame_no_grp]
                  xml.tag("xdr:cNvGraphicFramePr") { xml.empty_tag("a:graphicFrameLocks", { noGrp: 1 }) }
                else
                  xml.empty_tag("xdr:cNvGraphicFramePr")
                end
              end
              xml.tag("xdr:xfrm") do
                xml.empty_tag("a:off", { x: 0, y: 0 })
                xml.empty_tag("a:ext", { cx: 5_000_000, cy: 3_000_000 })
              end
              xml.tag("a:graphic") do
                xml.tag("a:graphicData", { uri: C_NS }) do
                  xml.empty_tag("c:chart", { "xmlns:c": C_NS, "xmlns:r": DOC_REL_NS, "r:id": rid })
                end
              end
            end
            write_client_data_xml(xml, chart)
          end
        end

        # : (XmlBuilder xml, untyped part) -> void
        def write_shape_anchor_xml(xml, part)
          shape = part[:shape]
          shape = shape.to_h if shape.is_a?(Elements::MetadataRecord)
          xml.tag("xdr:twoCellAnchor", { editAs: shape[:edit_as], fPublished: (1 if shape[:published]) }) do
            write_anchor_xml(xml, "from", shape[:from_column] || shape[:from_col], shape[:from_row], col_off: shape[:from_col_off] || 0, row_off: shape[:from_row_off] || 0)
            write_anchor_xml(xml, "to", shape[:to_column] || shape[:to_col], shape[:to_row], col_off: shape[:to_col_off] || 0, row_off: shape[:to_row_off] || 0)
            xml.tag("xdr:sp", { macro: shape[:macro], textlink: shape[:textlink] }) do
              xml.tag("xdr:nvSpPr") do
                xml.empty_tag(
                  "xdr:cNvPr",
                  {
                    id: part[:id] || "",
                    name: shape[:name] || "",
                    descr: shape[:description],
                    title: shape[:title],
                    hidden: (1 if shape[:hidden])
                  }
                )
                if shape[:no_grp] || shape[:no_rot] || shape[:f_locks_text]
                  xml.tag("xdr:cNvSpPr") do
                    xml.empty_tag(
                      "a:spLocks",
                      {
                        noGrp: (1 if shape[:no_grp]),
                        noRot: (1 if shape[:no_rot]),
                        fLocksText: (1 if shape[:f_locks_text])
                      }
                    )
                  end
                else
                  xml.empty_tag("xdr:cNvSpPr")
                end
              end
              xml.tag("xdr:spPr") do
                if shape[:rotation]
                  xml.tag("a:xfrm", { rot: shape[:rotation].to_i }) do
                    xml.empty_tag("a:off", { x: 0, y: 0 })
                    xml.empty_tag("a:ext", { cx: 0, cy: 0 })
                  end
                end
                xml.tag("a:prstGeom", { prst: shape[:preset] || "" }) do
                  if shape[:adjust_values]&.any?
                    xml.tag("a:avLst") do
                      shape[:adjust_values].each do |gd|
                        xml.empty_tag("a:gd", { name: gd[:name] || "", fmla: gd[:fmla] || "" })
                      end
                    end
                  else
                    xml.empty_tag("a:avLst")
                  end
                end
                write_shape_fill_xml(xml, shape)
                write_shape_line_xml(xml, shape)
                write_effect_lst_xml(xml, shape)
              end
              write_shape_text_body_xml(xml, shape) if shape[:text] || shape[:text_paragraphs]
            end
            write_client_data_xml(xml, shape)
          end
        end

        # : (XmlBuilder xml, untyped shape) -> void
        def write_shape_fill_xml(xml, shape)
          if shape[:no_fill]
            xml.empty_tag("a:noFill")
          elsif shape[:fill_color]
            xml.tag("a:solidFill") do
              write_drawing_color_xml(xml, shape[:fill_color], alpha: shape[:fill_alpha], transforms: shape[:fill_color_transforms])
            end
          elsif (gf = shape[:gradient_fill])
            rot = xml_bool_attr(gf[:rot_with_shape])
            xml.tag("a:gradFill", { rotWithShape: rot, flip: gf[:flip] }) do
              xml.tag("a:gsLst") do
                (gf[:stops] || []).each do |gs|
                  xml.tag("a:gs", { pos: gs[:pos] || "" }) { write_drawing_color_xml(xml, gs[:color]) }
                end
              end
              if gf[:path]
                xml.empty_tag("a:path", { path: gf[:path] })
              elsif gf[:angle]
                xml.empty_tag("a:lin", { ang: gf[:angle], scaled: gf[:scaled] ? 1 : 0 })
              end
              xml.empty_tag("a:tileRect", gf[:tile_rect]) if gf[:tile_rect]
            end
          elsif (pf = shape[:pattern_fill])
            xml.tag("a:pattFill", { prst: pf[:preset] || "" }) do
              xml.tag("a:fgClr") { write_drawing_color_xml(xml, pf[:fg_color]) } if pf[:fg_color]
              xml.tag("a:bgClr") { write_drawing_color_xml(xml, pf[:bg_color]) } if pf[:bg_color]
            end
          end
        end

        # : (XmlBuilder xml, untyped shape) -> void
        def write_shape_line_xml(xml, shape)
          if shape[:no_line]
            xml.tag("a:ln") { xml.empty_tag("a:noFill") }
          elsif shape[:line_color] || shape[:line_dash] || shape[:line_custom_dash] || shape[:head_end] || shape[:tail_end] || shape[:line_cap] || shape[:line_align] || shape[:line_compound] || shape[:line_join]
            xml.tag(
              "a:ln",
              {
                w: shape[:line_width]&.to_i,
                cap: shape[:line_cap],
                algn: shape[:line_align],
                cmpd: shape[:line_compound]
              }
            ) do
              if shape[:line_color]
                xml.tag("a:solidFill") do
                  write_drawing_color_xml(xml, shape[:line_color], alpha: shape[:line_alpha], transforms: shape[:line_color_transforms])
                end
              end
              if shape[:line_custom_dash]
                xml.tag("a:custDash") do
                  shape[:line_custom_dash].each do |ds|
                    xml.empty_tag("a:ds", { d: ds[:d] || "", sp: ds[:sp] || "" })
                  end
                end
              elsif shape[:line_dash]
                xml.empty_tag("a:prstDash", { val: shape[:line_dash] })
              end
              write_line_join_xml(xml, shape[:line_join], shape[:line_miter_limit])
              write_line_end_xml(xml, "headEnd", shape[:head_end])
              write_line_end_xml(xml, "tailEnd", shape[:tail_end])
            end
          end
        end

        # : (XmlBuilder xml, untyped line_join, untyped miter_limit) -> void
        def write_line_join_xml(xml, line_join, miter_limit)
          case line_join
          when "round" then xml.empty_tag("a:round")
          when "bevel" then xml.empty_tag("a:bevel")
          when "miter" then xml.empty_tag("a:miter", { lim: miter_limit&.to_i })
          end
        end

        # : (XmlBuilder xml, untyped source) -> void
        def write_effect_lst_xml(xml, source)
          return unless source[:outer_shadow] || source[:inner_shadow] || source[:glow] || source[:soft_edge] || source[:reflection] || source[:blur]

          xml.tag("a:effectLst") do
            write_shadow_xml(xml, "a:outerShdw", source[:outer_shadow], include_algn: true) if source[:outer_shadow]
            if (gl = source[:glow])
              xml.tag("a:glow", { rad: gl[:rad] }) do
                write_drawing_color_xml(xml, gl[:color], alpha: gl[:alpha]) if gl[:color]
              end
            end
            write_shadow_xml(xml, "a:innerShdw", source[:inner_shadow], include_algn: false) if source[:inner_shadow]
            if (rf = source[:reflection])
              xml.empty_tag(
                "a:reflection",
                {
                  blurRad: rf[:blur_rad],
                  stA: rf[:st_a],
                  endA: rf[:end_a],
                  dist: rf[:dist],
                  dir: rf[:dir],
                  fadeDir: rf[:fade_dir],
                  sx: rf[:sx],
                  sy: rf[:sy],
                  kx: rf[:kx],
                  ky: rf[:ky],
                  algn: rf[:algn],
                  rotWithShape: xml_bool_attr(rf[:rot_with_shape])
                }
              )
            end
            xml.empty_tag("a:softEdge", { rad: source[:soft_edge][:rad] || "" }) if source[:soft_edge]
            if (bl = source[:blur])
              xml.empty_tag(
                "a:blur",
                {
                  rad: bl[:rad],
                  grow: xml_bool_attr(bl[:grow])
                }
              )
            end
          end
        end

        # : (XmlBuilder xml, ::String tag, untyped spec, include_algn: bool) -> void
        def write_shadow_xml(xml, tag, spec, include_algn:)
          attrs = { blurRad: spec[:blur_rad], dist: spec[:dist], dir: spec[:dir] }
          if include_algn
            attrs[:algn] = spec[:algn]
            attrs[:rotWithShape] = (spec[:rot_with_shape] ? 1 : 0) unless spec[:rot_with_shape].nil?
          end
          xml.tag(tag, attrs) do
            write_drawing_color_xml(xml, spec[:color], alpha: spec[:alpha]) if spec[:color]
          end
        end

        # : (XmlBuilder xml, untyped shape) -> void
        def write_shape_text_body_xml(xml, shape)
          ins = shape[:text_insets]
          body_pr_attrs = {
            rot: shape[:text_rot],
            spcFirstLastPara: (1 if shape[:text_spc_first_last_para]),
            wrap: shape[:text_wrap],
            anchor: shape[:text_anchor],
            anchorCtr: (1 if shape[:text_anchor_ctr]),
            vertOverflow: shape[:text_vert_overflow],
            horzOverflow: shape[:text_horz_overflow],
            numCol: shape[:text_num_col],
            spcCol: shape[:text_spc_col],
            rtlCol: (1 if shape[:text_rtl_col]),
            fromWordArt: (1 if shape[:text_from_word_art]),
            upright: (1 if shape[:text_upright]),
            compatLnSpc: (1 if shape[:text_compat_ln_spc]),
            forceAA: (1 if shape[:text_force_aa]),
            vert: shape[:text_vertical],
            lIns: ins&.[](:left),
            tIns: ins&.[](:top),
            rIns: ins&.[](:right),
            bIns: ins&.[](:bottom)
          }
          autofit = shape[:autofit]
          has_autofit = %w[none shape normal].include?(autofit) || (autofit.is_a?(Hash) && %w[normal none shape].include?(autofit[:type]))
          xml.tag("xdr:txBody") do
            if shape[:text_warp] || has_autofit
              xml.tag("a:bodyPr", body_pr_attrs) do
                if (tw = shape[:text_warp])
                  xml.tag("a:prstTxWarp", { prst: tw[:preset] || "" }) { xml.empty_tag("a:avLst") }
                end
                write_autofit_xml(xml, autofit) if has_autofit
              end
            else
              xml.empty_tag("a:bodyPr", body_pr_attrs)
            end
            xml.empty_tag("a:lstStyle")
            if shape[:text_paragraphs]
              shape[:text_paragraphs].each { |para| write_paragraph_xml(xml, para) }
            elsif shape[:text]
              write_paragraph_xml(
                xml,
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
          end
        end

        # : (XmlBuilder xml, untyped autofit) -> void
        def write_autofit_xml(xml, autofit)
          case autofit
          when "none" then xml.empty_tag("a:noAutofit")
          when "shape" then xml.empty_tag("a:spAutoFit")
          when "normal" then xml.empty_tag("a:normAutofit")
          when Hash
            case autofit[:type]
            when "normal"
              xml.empty_tag("a:normAutofit", { fontScale: autofit[:font_scale], lnSpcReduction: autofit[:ln_spc_reduction] })
            when "none"
              xml.empty_tag("a:noAutofit")
            when "shape"
              xml.empty_tag("a:spAutoFit")
            end
          end
        end

        # : (XmlBuilder xml, untyped para) -> void
        def write_paragraph_xml(xml, para)
          ti = para[:indent]
          ppr_attrs = {
            algn: para[:align],
            fontAlgn: para[:font_align],
            defTabSz: para[:def_tab_sz],
            rtl: (1 if para[:rtl]),
            eaLnBrk: (1 if para[:ea_ln_brk]),
            latinLnBrk: (1 if para[:latin_ln_brk]),
            hangingPunct: (1 if para[:hanging_punct]),
            lvl: para[:level],
            marL: ti&.[](:left),
            marR: ti&.[](:right),
            indent: ti&.[](:indent)
          }.compact
          has_children = paragraph_ppr_has_children?(para)

          xml.tag("a:p") do
            if has_children
              xml.tag("a:pPr", ppr_attrs) { write_paragraph_ppr_children_xml(xml, para) }
            elsif !ppr_attrs.empty?
              xml.empty_tag("a:pPr", ppr_attrs)
            end
            if para[:runs]
              para[:runs].each do |run|
                r_font = run.respond_to?(:font) ? run.font : run[:font]
                r_text = run.respond_to?(:text) ? run.text : run[:text]
                xml.tag("a:r") do
                  write_text_char_props_xml(xml, "a:rPr", r_font) if r_font
                  xml.text_tag("a:t", r_text || "")
                end
              end
            else
              xml.tag("a:r") do
                write_text_char_props_xml(xml, "a:rPr", para[:font]) if para[:font]
                xml.text_tag("a:t", para[:text] || "")
              end
            end
            write_text_char_props_xml(xml, "a:endParaRPr", para[:end_para_rpr]) if para[:end_para_rpr]
          end
        end

        # : (untyped para) -> bool
        def paragraph_ppr_has_children?(para)
          return true if para[:tab_stops]&.any? || para[:def_rpr]

          ts = para[:spacing]
          return true if ts && (ts[:line] || ts[:line_pct] || ts[:before] || ts[:before_pct] || ts[:after] || ts[:after_pct])

          bu = para[:bullet]
          return true if bu && (bu[:color] || bu[:size_pts] || bu[:size_pct] || bu[:font] || %w[none char auto].include?(bu[:type]))

          false
        end

        # : (XmlBuilder xml, untyped para) -> void
        def write_paragraph_ppr_children_xml(xml, para)
          if (ts = para[:spacing])
            if ts[:line]
              xml.tag("a:lnSpc") { xml.empty_tag("a:spcPts", { val: ts[:line] }) }
            elsif ts[:line_pct]
              xml.tag("a:lnSpc") { xml.empty_tag("a:spcPct", { val: ts[:line_pct] }) }
            end
            if ts[:before]
              xml.tag("a:spcBef") { xml.empty_tag("a:spcPts", { val: ts[:before] }) }
            elsif ts[:before_pct]
              xml.tag("a:spcBef") { xml.empty_tag("a:spcPct", { val: ts[:before_pct] }) }
            end
            if ts[:after]
              xml.tag("a:spcAft") { xml.empty_tag("a:spcPts", { val: ts[:after] }) }
            elsif ts[:after_pct]
              xml.tag("a:spcAft") { xml.empty_tag("a:spcPct", { val: ts[:after_pct] }) }
            end
          end
          if para[:tab_stops]&.any?
            xml.tag("a:tabLst") do
              para[:tab_stops].each do |tab|
                xml.empty_tag("a:tab", { pos: tab[:pos] || "", algn: tab[:align] })
              end
            end
          end
          if (bu = para[:bullet])
            xml.tag("a:buClr") { write_drawing_color_xml(xml, bu[:color]) } if bu[:color]
            xml.empty_tag("a:buSzPts", { val: bu[:size_pts] }) if bu[:size_pts]
            xml.empty_tag("a:buSzPct", { val: bu[:size_pct] }) if bu[:size_pct]
            xml.empty_tag("a:buFont", { typeface: bu[:font] }) if bu[:font]
            case bu[:type]
            when "none"
              xml.empty_tag("a:buNone")
            when "char"
              xml.empty_tag("a:buChar", { char: bu[:char] || "" })
            when "auto"
              xml.empty_tag("a:buAutoNum", { type: bu[:auto_type] || "", startAt: bu[:start_at] })
            end
          end
          write_text_char_props_xml(xml, "a:defRPr", para[:def_rpr]) if para[:def_rpr]
        end

        # : (XmlBuilder xml, ::String tag, untyped font) -> void
        def write_text_char_props_xml(xml, tag, font)
          attrs = {
            b: (1 if font[:bold]),
            i: (1 if font[:italic]),
            noProof: (1 if font[:no_proof]),
            normalizeH: (1 if font[:normalize_h]),
            kumimoji: (1 if font[:kumimoji]),
            sz: font[:size],
            strike: font[:strike],
            u: font[:underline],
            baseline: font[:baseline],
            spc: font[:spacing],
            kern: font[:kern],
            cap: font[:cap],
            lang: font[:lang],
            altLang: font[:alt_lang],
            dirty: (1 if font[:dirty]),
            smtClean: (1 if font[:smt_clean]),
            err: (1 if font[:err]),
            bmk: font[:bmk]
          }
          has_ln = font[:line_color] || font[:line_width] || font[:line_dash] || font[:line_cap] || font[:line_join]
          has_eff = font[:outer_shadow] || font[:inner_shadow] || font[:glow] || font[:soft_edge] || font[:reflection] || font[:blur]
          has_children = has_ln || font[:color] || has_eff || font[:highlight] || font[:u_ln_tx] || font[:u_fill_tx] || font[:name] || font[:ea_font] || font[:cs_font] || font[:sym_font]

          unless has_children
            xml.empty_tag(tag, attrs)
            return
          end

          xml.tag(tag, attrs) do
            if has_ln
              xml.tag("a:ln", { w: font[:line_width]&.to_i, cap: font[:line_cap] }) do
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, font[:line_color]) } if font[:line_color]
                xml.empty_tag("a:prstDash", { val: font[:line_dash] }) if font[:line_dash]
                write_line_join_xml(xml, font[:line_join], font[:line_miter_limit])
              end
            end
            xml.tag("a:solidFill") { write_drawing_color_xml(xml, font[:color]) } if font[:color]
            write_effect_lst_xml(xml, font)
            xml.tag("a:highlight") { write_drawing_color_xml(xml, font[:highlight]) } if font[:highlight]
            xml.empty_tag("a:uLnTx") if font[:u_ln_tx]
            xml.empty_tag("a:uFillTx") if font[:u_fill_tx]
            xml.empty_tag("a:latin", { typeface: font[:name] }) if font[:name]
            xml.empty_tag("a:ea", { typeface: font[:ea_font] }) if font[:ea_font]
            xml.empty_tag("a:cs", { typeface: font[:cs_font] }) if font[:cs_font]
            xml.empty_tag("a:sym", { typeface: font[:sym_font] }) if font[:sym_font]
          end
        end

        # : (XmlBuilder xml, ::String tag, untyped col, untyped row, ?col_off: ::Integer, ?row_off: ::Integer) -> void
        def write_anchor_xml(xml, tag, col, row, col_off: 0, row_off: 0)
          c = col.to_s.empty? ? 0 : col.to_i
          r = row.to_s.empty? ? 0 : row.to_i
          co = col_off.to_s.empty? ? 0 : col_off.to_i
          ro = row_off.to_s.empty? ? 0 : row_off.to_i
          xml.tag("xdr:#{tag}") do
            xml.text_tag("xdr:col", c)
            xml.text_tag("xdr:colOff", co)
            xml.text_tag("xdr:row", r)
            xml.text_tag("xdr:rowOff", ro)
          end
        end

        # : (XmlBuilder xml, untyped obj) -> void
        def write_client_data_xml(xml, obj)
          locks = case obj[:locks_with_sheet]
                  when false then 0
                  when true then 1
                  end
          prints = case obj[:prints_with_sheet]
                   when false then 0
                   when true then 1
                   end
          xml.empty_tag("xdr:clientData", { fLocksWithSheet: locks, fPrintsWithSheet: prints })
        end

        # : (XmlBuilder xml, untyped title_spec, ?overlay: bool) -> void
        def write_chart_title_xml(xml, title_spec, overlay: false)
          overlay_val = overlay ? 1 : 0
          xml.tag("c:title") do
            if title_spec.is_a?(Hash)
              text = title_spec[:text].to_s
              rotation = title_spec[:rotation]
              font = title_spec[:font] || {}
              xml.tag("c:tx") do
                xml.tag("c:rich") do
                  xml.empty_tag("a:bodyPr", { rot: rotation })
                  xml.empty_tag("a:lstStyle")
                  xml.tag("a:p") do
                    xml.tag("a:r") do
                      rpr_attrs = { b: (1 if font[:bold]), i: (1 if font[:italic]), sz: font[:size] }
                      if font[:color] || font[:name]
                        xml.tag("a:rPr", rpr_attrs) do
                          xml.tag("a:solidFill") { write_drawing_color_xml(xml, font[:color]) } if font[:color]
                          xml.empty_tag("a:latin", { typeface: font[:name] }) if font[:name]
                        end
                      else
                        xml.empty_tag("a:rPr", rpr_attrs)
                      end
                      xml.text_tag("a:t", text)
                    end
                  end
                end
              end
              write_title_layout_xml(xml, title_spec[:layout])
              xml.empty_tag("c:overlay", { val: overlay_val })
              write_title_sp_pr(xml, title_spec)
            else
              xml.tag("c:tx") do
                xml.tag("c:rich") do
                  xml.empty_tag("a:bodyPr")
                  xml.empty_tag("a:lstStyle")
                  xml.tag("a:p") { xml.tag("a:r") { xml.text_tag("a:t", title_spec || "") } }
                end
              end
              xml.empty_tag("c:overlay", { val: overlay_val })
            end
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

        # : (XmlBuilder xml, untyped spec) -> void
        def write_title_sp_pr(xml, spec)
          write_simple_sp_pr_from_hash(xml, spec)
        end

        # : (XmlBuilder xml, untyped layout) -> void
        def write_title_layout_xml(xml, layout)
          return unless layout.is_a?(Hash)

          attrs = { x: layout[:x], y: layout[:y], w: layout[:w], h: layout[:h] }.compact
          return if attrs.empty?

          xml.tag("c:layout") do
            xml.tag("c:manualLayout") do
              attrs.each { |k, v| xml.empty_tag("c:#{k}", { val: v }) }
            end
          end
        end

        # : (XmlBuilder xml, ::String tag, untyped spec) -> void
        def write_gridlines_xml(xml, tag, spec)
          write_line_sp_tag(xml, "c:#{tag}", spec)
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

        # : (XmlBuilder xml, untyped ref) -> void
        def write_num_cache_xml(xml, ref)
          values = resolve_sheet_ref(ref)
          return unless values

          xml.tag("c:numCache") do
            xml.text_tag("c:formatCode", "General")
            xml.empty_tag("c:ptCount", { val: values.size })
            values.each_with_index do |v, i|
              next unless v

              xml.tag("c:pt", { idx: i }) { xml.text_tag("c:v", v.to_s) }
            end
          end
        end

        # : (XmlBuilder xml, untyped ref) -> void
        def write_str_cache_xml(xml, ref)
          values = resolve_sheet_ref(ref)
          return unless values

          xml.tag("c:strCache") do
            xml.empty_tag("c:ptCount", { val: values.size })
            values.each_with_index do |v, i|
              next unless v

              xml.tag("c:pt", { idx: i }) { xml.text_tag("c:v", v.to_s) }
            end
          end
        end

        # : (XmlBuilder xml, untyped color, ?alpha: untyped?, ?transforms: untyped?) -> void
        def write_drawing_color_xml(xml, color, alpha: nil, transforms: nil)
          if color.is_a?(Hash) && color[:scheme]
            tag = "a:schemeClr"
            val = color[:scheme]
            transforms = color[:transforms] || transforms
          else
            tag = "a:srgbClr"
            val = normalize_drawing_rgb(color)
          end
          has_children = alpha || (transforms && !transforms.empty?)
          unless has_children
            xml.empty_tag(tag, { val: val || "" })
            return
          end

          xml.tag(tag, { val: val || "" }) do
            xml.empty_tag("a:alpha", { val: alpha }) if alpha
            (transforms || []).each do |t|
              type = t[:type].to_s
              raise ArgumentError, "Unknown color transform: #{type.inspect}. Use one of: #{COLOR_TRANSFORMS.join(", ")}." unless COLOR_TRANSFORMS.include?(type)

              xml.empty_tag("a:#{type}", { val: t[:val] || "" })
            end
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

        # : (XmlBuilder xml, untyped rotation, untyped font) -> void
        def write_axis_txpr(xml, rotation, font)
          return unless rotation || font

          xml.tag("c:txPr") do
            xml.empty_tag("a:bodyPr", { rot: rotation })
            xml.empty_tag("a:lstStyle")
            xml.tag("a:p") do
              xml.tag("a:pPr") do
                rpr_attrs = if font
                              {
                                sz: (font[:size] ? (font[:size] * 100).to_i : nil),
                                b: (1 if font[:bold]),
                                i: (1 if font[:italic])
                              }
                            else
                              XmlBuilder::EMPTY_HASH
                            end
                if font && (font[:color] || font[:name])
                  xml.tag("a:defRPr", rpr_attrs) do
                    xml.tag("a:solidFill") { write_drawing_color_xml(xml, font[:color]) } if font[:color]
                    xml.empty_tag("a:latin", { typeface: font[:name] }) if font[:name]
                  end
                else
                  xml.empty_tag("a:defRPr", rpr_attrs)
                end
              end
              xml.empty_tag("a:endParaRPr")
            end
          end
        end

        # : (XmlBuilder xml, ax_id: ::Integer, prefix: ::Symbol, chart: untyped, ?default_pos: ::String, ?default_crosses: ::String?) -> void
        def write_val_ax_xml(xml, ax_id:, prefix:, chart:, default_pos: "l", default_crosses: nil)
          del_key = :"#{prefix}_delete"
          orient_key = :"#{prefix}_orientation"
          log_key = :"#{prefix}_log_base"
          max_key = :"#{prefix}_scaling_max"
          min_key = :"#{prefix}_scaling_min"
          pos_key = :"#{prefix}_pos"
          major_grid_key = :"#{prefix}_major_gridlines"
          minor_grid_key = :"#{prefix}_minor_gridlines"
          title_key = :"#{prefix}_title"
          num_fmt_key = :"#{prefix}_num_fmt"
          major_tick_key = :"#{prefix}_major_tick_mark"
          minor_tick_key = :"#{prefix}_minor_tick_mark"
          tick_lbl_pos_key = :"#{prefix}_tick_lbl_pos"
          line_color_key = :"#{prefix}_line_color"
          fill_key = :"#{prefix}_fill"
          no_fill_key = :"#{prefix}_no_fill"
          line_dash_key = :"#{prefix}_line_dash"
          line_width_key = :"#{prefix}_line_width"
          rotation_key = :"#{prefix}_label_rotation"
          font_key = :"#{prefix}_font"
          crosses_key = :"#{prefix}_crosses"
          crosses_at_key = :"#{prefix}_crosses_at"
          cross_between_key = :"#{prefix}_cross_between"
          major_unit_key = :"#{prefix}_major_unit"
          minor_unit_key = :"#{prefix}_minor_unit"
          disp_units_key = :"#{prefix}_disp_units"

          xml.tag("c:valAx") do
            xml.empty_tag("c:axId", { val: ax_id })
            xml.tag("c:scaling") do
              xml.empty_tag("c:logBase", { val: chart[log_key] }) if chart[log_key]
              xml.empty_tag("c:orientation", { val: chart[orient_key] || "minMax" })
              xml.empty_tag("c:max", { val: chart[max_key] }) if chart[max_key]
              xml.empty_tag("c:min", { val: chart[min_key] }) if chart[min_key]
            end
            xml.empty_tag("c:delete", { val: chart[del_key] ? 1 : 0 })
            xml.empty_tag("c:axPos", { val: chart[pos_key] || default_pos })
            write_gridlines_xml(xml, "majorGridlines", chart[major_grid_key]) if chart[major_grid_key]
            write_gridlines_xml(xml, "minorGridlines", chart[minor_grid_key]) if chart[minor_grid_key]
            if chart[title_key]
              val_title_spec = merge_flat_title_styling(chart[title_key], chart, title_key)
              write_chart_title_xml(xml, val_title_spec)
            end
            if (vnf = chart[num_fmt_key])
              vnf_hash = vnf.is_a?(Hash) ? vnf : { format_code: vnf }
              xml.empty_tag("c:numFmt", { formatCode: vnf_hash[:format_code] || "", sourceLinked: vnf_hash[:source_linked] ? 1 : 0 })
            end
            xml.empty_tag("c:majorTickMark", { val: chart[major_tick_key] }) if chart[major_tick_key]
            xml.empty_tag("c:minorTickMark", { val: chart[minor_tick_key] }) if chart[minor_tick_key]
            xml.empty_tag("c:tickLblPos", { val: chart[tick_lbl_pos_key] }) if chart[tick_lbl_pos_key]
            if chart[line_color_key] || chart[fill_key] || chart[no_fill_key] || chart[line_dash_key]
              xml.tag("c:spPr") do
                xml.empty_tag("a:noFill") if chart[no_fill_key]
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, chart[fill_key]) } if chart[fill_key]
                if chart[line_color_key] || chart[line_dash_key]
                  xml.tag("a:ln", { w: chart[line_width_key] }) do
                    xml.tag("a:solidFill") { write_drawing_color_xml(xml, chart[line_color_key]) } if chart[line_color_key]
                    xml.empty_tag("a:prstDash", { val: chart[line_dash_key] }) if chart[line_dash_key]
                  end
                end
              end
            end
            write_axis_txpr(xml, chart[rotation_key], chart[font_key])
            xml.empty_tag("c:crossAx", { val: 1 })
            crosses_val = chart[crosses_key] || default_crosses
            xml.empty_tag("c:crosses", { val: crosses_val }) if crosses_val
            xml.empty_tag("c:crossesAt", { val: chart[crosses_at_key] }) if !crosses_val && chart[crosses_at_key]
            xml.empty_tag("c:crossBetween", { val: chart[cross_between_key] }) if chart[cross_between_key]
            xml.empty_tag("c:majorUnit", { val: chart[major_unit_key] }) if chart[major_unit_key]
            xml.empty_tag("c:minorUnit", { val: chart[minor_unit_key] }) if chart[minor_unit_key]
            write_disp_units_xml(xml, chart[disp_units_key]) if chart[disp_units_key]
          end
        end

        # : (XmlBuilder xml, untyped disp_units) -> void
        def write_disp_units_xml(xml, disp_units)
          xml.tag("c:dispUnits") do
            if disp_units.is_a?(Hash)
              xml.empty_tag("c:builtInUnit", { val: disp_units[:built_in_unit] }) if disp_units[:built_in_unit]
              xml.empty_tag("c:custUnit", { val: disp_units[:cust_unit] }) if disp_units[:cust_unit]
              if (dul = disp_units[:label])
                xml.tag("c:dispUnitsLbl") do
                  if dul[:num_fmt]
                    nf = dul[:num_fmt]
                    nf_src = nf.is_a?(Hash) ? nf : { format_code: nf }
                    nf_linked = nf_src[:source_linked] ? 1 : 0
                    xml.empty_tag("c:numFmt", { formatCode: nf_src[:format_code] || "", sourceLinked: nf_linked })
                  end
                  write_simple_sp_pr_from_hash(xml, dul)
                  write_axis_txpr(xml, nil, dul[:font]) if dul[:font]
                end
              end
            else
              xml.empty_tag("c:builtInUnit", { val: disp_units.to_s })
            end
          end
        end

        # : (XmlBuilder xml, untyped chart) -> void
        def write_chart_top_title_xml(xml, chart)
          return unless chart[:title]

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
          write_chart_title_xml(xml, title_spec, overlay: chart[:title_overlay])
        end

        # : (XmlBuilder xml, untyped v3d) -> void
        def write_view_3d_xml(xml, v3d)
          xml.tag("c:view3D") do
            xml.empty_tag("c:rotX", { val: v3d[:rot_x] }) if v3d[:rot_x]
            xml.empty_tag("c:hPercent", { val: v3d[:h_percent] }) if v3d[:h_percent]
            xml.empty_tag("c:rotY", { val: v3d[:rot_y] }) if v3d[:rot_y]
            xml.empty_tag("c:depthPercent", { val: v3d[:depth_percent] }) if v3d[:depth_percent]
            xml.empty_tag("c:rAngAx", { val: v3d[:r_ang_ax] ? 1 : 0 }) unless v3d[:r_ang_ax].nil?
            xml.empty_tag("c:perspective", { val: v3d[:perspective] }) if v3d[:perspective]
          end
        end

        # : (XmlBuilder xml, untyped chart) -> void
        def write_chart_walls_xml(xml, chart)
          { floor: "floor", side_wall: "sideWall", back_wall: "backWall" }.each do |wall_key, tag|
            next unless (wall = chart[wall_key])

            xml.tag("c:#{tag}") do
              xml.tag("c:spPr") do
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, wall[:fill_color]) } if wall[:fill_color]
                xml.empty_tag("a:noFill") if wall[:no_fill]
                if wall[:line_color] || wall[:line_dash]
                  xml.tag("a:ln", { w: wall[:line_width] }) do
                    xml.tag("a:solidFill") { write_drawing_color_xml(xml, wall[:line_color]) } if wall[:line_color]
                    xml.empty_tag("a:prstDash", { val: wall[:line_dash] }) if wall[:line_dash]
                  end
                end
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped chart, ::String chart_type) -> void
        def write_plot_area_xml(xml, chart, chart_type)
          xml.tag("c:plotArea") do
            write_plot_area_layout_xml(xml, chart[:plot_area_layout])
            all_series = chart[:series] || []
            all_series.each do |s|
              s[:cat_ref] ||= s[:categories] if s[:categories]
              s[:val_ref] ||= s[:values] if s[:values]
            end

            has_secondary = !chart[:sec_val_axis_title].nil? ||
                            !chart[:secondary_value_axis].nil? ||
                            all_series.any? { |s| s[:secondary_axis] || s[:axis] == :secondary || s[:axis] == "secondary" }

            groups = group_chart_series(all_series, chart_type)
            groups.each { |grp| write_chart_group_xml(xml, grp, chart) }

            write_chart_axes_xml(xml, chart, groups, has_secondary) if groups.any? { |g| !NO_AXIS_CHARTS.include?(g[:type]) }
            write_chart_data_table_xml(xml, chart[:data_table]) if chart[:data_table]

            if chart[:plot_area_fill] || chart[:plot_area_no_fill] || chart[:plot_area_line_color] || chart[:plot_area_line_dash]
              write_simple_sp_pr(
                xml,
                fill_color: chart[:plot_area_fill],
                no_fill: chart[:plot_area_no_fill],
                line_color: chart[:plot_area_line_color],
                line_width: chart[:plot_area_line_width],
                line_dash: chart[:plot_area_line_dash],
                emu: false
              )
            end
          end
        end

        # : (XmlBuilder xml, untyped layout, ?emit_empty: bool) -> void
        def write_plot_area_layout_xml(xml, layout, emit_empty: true)
          ml_entries = if layout.is_a?(Hash)
                         [
                           ["c:layoutTarget", layout[:target]],
                           ["c:xMode", layout[:x_mode]],
                           ["c:yMode", layout[:y_mode]],
                           ["c:wMode", layout[:w_mode]],
                           ["c:hMode", layout[:h_mode]],
                           ["c:x", layout[:x]],
                           ["c:y", layout[:y]],
                           ["c:w", layout[:w]],
                           ["c:h", layout[:h]]
                         ].select { |_, val| val }
                       else
                         []
                       end
          if ml_entries.empty?
            xml.empty_tag("c:layout") if emit_empty
          else
            xml.tag("c:layout") do
              xml.tag("c:manualLayout") do
                ml_entries.each { |tag_name, val| xml.empty_tag(tag_name, { val: val }) }
              end
            end
          end
        end

        # : (untyped all_series, ::String chart_type) -> ::Array[untyped]
        def group_chart_series(all_series, chart_type)
          groups = []
          all_series.each_with_index do |ser, idx|
            s_type = if ser[:type]
                       # The type is written as an element name, so only known types are accepted.
                       CHART_TYPE_MAP[ser[:type]] || CHART_TYPE_MAP.values.find { |element| element == ser[:type].to_s } ||
                         raise(ArgumentError, "Unknown chart series type: #{ser[:type].inspect}. Use one of: #{CHART_TYPE_MAP.keys.map(&:inspect).join(", ")}.")
                     elsif chart_type == "comboChart"
                       "barChart"
                     else
                       chart_type
                     end
            s_sec = ser[:secondary_axis] || ser[:axis] == :secondary || ser[:axis] == "secondary"
            s_val_ax = s_sec ? 3 : 2

            group = groups.find { |g| g[:type] == s_type && g[:val_ax_id] == s_val_ax }
            if group
              group[:series] << [ser, idx]
            else
              groups << { type: s_type, val_ax_id: s_val_ax, series: [[ser, idx]] }
            end
          end

          if groups.empty?
            groups << { type: (chart_type == "comboChart" ? "barChart" : chart_type), val_ax_id: 2, series: [] }
          end
          groups
        end

        # : (XmlBuilder xml, untyped grp, untyped chart) -> void
        def write_chart_group_xml(xml, grp, chart)
          g_type = grp[:type]
          g_val_ax = grp[:val_ax_id]
          g_no_axes = NO_AXIS_CHARTS.include?(g_type)

          xml.tag("c:#{g_type}") do
            xml.empty_tag("c:ofPieType", { val: chart[:of_pie_type] || "pie" }) if g_type == "ofPieChart"
            if %w[barChart bar3DChart].include?(g_type)
              xml.empty_tag("c:barDir", { val: chart[:bar_dir] || "col" })
              xml.empty_tag("c:grouping", { val: chart[:grouping] || "clustered" })
            elsif GROUPING_CHARTS.include?(g_type)
              xml.empty_tag("c:grouping", { val: chart[:grouping] || "standard" })
            end
            xml.empty_tag("c:scatterStyle", { val: chart[:scatter_style] || "lineMarker" }) if g_type == "scatterChart"
            xml.empty_tag("c:radarStyle", { val: chart[:radar_style] || "standard" }) if g_type == "radarChart"
            xml.empty_tag("c:varyColors", { val: chart[:vary_colors] ? 1 : 0 }) unless chart[:vary_colors].nil?
            xml.empty_tag("c:wireframe", { val: chart[:wireframe] ? 1 : 0 }) unless chart[:wireframe].nil?

            grp[:series].each { |ser, idx| write_chart_series_xml(xml, ser, idx, g_type, chart) }

            write_band_fmts_xml(xml, chart[:band_fmts]) if chart[:band_fmts]
            xml.empty_tag("c:gapWidth", { val: chart[:gap_width] }) if chart[:gap_width]
            xml.empty_tag("c:splitType", { val: chart[:split_type] }) if chart[:split_type]
            xml.empty_tag("c:splitPos", { val: chart[:split_pos] }) if chart[:split_pos]
            if chart[:cust_split]&.any?
              xml.tag("c:custSplit") do
                chart[:cust_split].each { |idx| xml.empty_tag("c:secondPiePt", { val: idx || "" }) }
              end
            end
            xml.empty_tag("c:secondPieSize", { val: chart[:second_pie_size] }) if chart[:second_pie_size]
            xml.empty_tag("c:gapDepth", { val: chart[:gap_depth] }) if chart[:gap_depth]
            xml.empty_tag("c:overlap", { val: chart[:overlap] }) if chart[:overlap]
            xml.empty_tag("c:shape", { val: chart[:bar_shape] }) if chart[:bar_shape]
            xml.empty_tag("c:bubble3D", { val: chart[:bubble_3d] ? 1 : 0 }) unless chart[:bubble_3d].nil?
            xml.empty_tag("c:bubbleScale", { val: chart[:bubble_scale] }) if chart[:bubble_scale]
            xml.empty_tag("c:showNegBubbles", { val: chart[:show_neg_bubbles] ? 1 : 0 }) unless chart[:show_neg_bubbles].nil?
            xml.empty_tag("c:sizeRepresents", { val: chart[:size_represents] }) if chart[:size_represents]
            # Doughnut charts carry both; Excel's defaults are 0 and 50.
            first_slice_ang = chart[:first_slice_ang] || (0 if g_type == "doughnutChart")
            hole_size = chart[:hole_size] || (50 if g_type == "doughnutChart")
            xml.empty_tag("c:firstSliceAng", { val: first_slice_ang }) if first_slice_ang
            xml.empty_tag("c:holeSize", { val: hole_size }) if hole_size
            write_line_sp_tag(xml, "c:serLines", chart[:ser_lines]) if chart[:ser_lines]
            write_line_sp_tag(xml, "c:dropLines", chart[:drop_lines]) if chart[:drop_lines]
            write_line_sp_tag(xml, "c:hiLowLines", chart[:hi_low_lines]) if chart[:hi_low_lines]
            write_up_down_bars_xml(xml, chart[:up_down_bars]) if chart[:up_down_bars]
            xml.empty_tag("c:marker", { val: chart[:marker] ? 1 : 0 }) unless chart[:marker].nil?
            xml.empty_tag("c:smooth", { val: chart[:smooth] ? 1 : 0 }) unless chart[:smooth].nil?
            unless g_no_axes
              xml.empty_tag("c:axId", { val: 1 })
              xml.empty_tag("c:axId", { val: g_val_ax })
              xml.empty_tag("c:axId", { val: 3 }) if THREE_AXIS_CHARTS.include?(g_type)
            end
          end
        end

        # : (XmlBuilder xml, untyped ser, ::Integer idx, ::String g_type, untyped chart) -> void
        def write_chart_series_xml(xml, ser, idx, g_type, chart)
          ser_order = ser[:order] || idx
          xml.tag("c:ser") do
            xml.empty_tag("c:idx", { val: idx })
            xml.empty_tag("c:order", { val: ser_order })
            if ser[:name]
              xml.tag("c:tx") do
                xml.tag("c:strRef") do
                  xml.text_tag("c:f", ser[:name])
                  write_str_cache_xml(xml, ser[:name])
                end
              end
            end

            # The series schemas order these children spPr, then invertIfNegative, marker, or
            # explosion (by chart type), then dPt.
            write_series_sp_pr_xml(xml, ser)
            xml.empty_tag("c:invertIfNegative", { val: ser[:invert_if_negative] ? 1 : 0 }) unless ser[:invert_if_negative].nil?
            xml.empty_tag("c:explosion", { val: ser[:explosion] }) if ser[:explosion]
            write_marker_xml(xml, ser)
            ser[:data_points]&.each { |dp| write_data_point_xml(xml, dp) }

            if (dl = ser[:data_labels] || chart[:data_labels])
              write_data_labels_xml(xml, dl)
            end
            trendline_list = ser[:trendlines] || (ser[:trendline] ? [ser[:trendline]] : [])
            trendline_list.each { |tl| write_trendline_xml(xml, tl) }
            err_bars_list = ser[:error_bars_list] || (ser[:error_bars] ? [ser[:error_bars]] : [])
            err_bars_list.each { |eb| write_error_bars_xml(xml, eb) }

            uses_xy = %w[scatterChart bubbleChart].include?(g_type)
            cat_tag = uses_xy ? "xVal" : "cat"
            val_tag = uses_xy ? "yVal" : "val"
            if ser[:cat_ref]
              xml.tag("c:#{cat_tag}") do
                if uses_xy || ser[:cat_ref_type] == :num
                  xml.tag("c:numRef") do
                    xml.text_tag("c:f", ser[:cat_ref])
                    write_num_cache_xml(xml, ser[:cat_ref])
                  end
                else
                  xml.tag("c:strRef") do
                    xml.text_tag("c:f", ser[:cat_ref])
                    write_str_cache_xml(xml, ser[:cat_ref])
                  end
                end
              end
            end
            if ser[:val_ref]
              xml.tag("c:#{val_tag}") do
                xml.tag("c:numRef") do
                  xml.text_tag("c:f", ser[:val_ref])
                  write_num_cache_xml(xml, ser[:val_ref])
                end
              end
            end
            if ser[:bubble_size_ref]
              xml.tag("c:bubbleSize") do
                xml.tag("c:numRef") do
                  xml.text_tag("c:f", ser[:bubble_size_ref])
                  write_num_cache_xml(xml, ser[:bubble_size_ref])
                end
              end
            end
            xml.empty_tag("c:smooth", { val: ser[:smooth] ? 1 : 0 }) unless ser[:smooth].nil?
            xml.empty_tag("c:shape", { val: ser[:shape] }) if ser[:shape]
          end
        end

        # : (XmlBuilder xml, untyped ser) -> void
        def write_series_sp_pr_xml(xml, ser)
          has_ln = ser[:line_color] || ser[:no_line] || ser[:line_width] || ser[:line_cap] || ser[:line_join] || ser[:line_dash]
          return unless ser[:fill_color] || ser[:no_fill] || has_ln

          xml.tag("c:spPr") do
            xml.tag("a:solidFill") { write_drawing_color_xml(xml, ser[:fill_color]) } if ser[:fill_color]
            xml.empty_tag("a:noFill") if ser[:no_fill]
            if has_ln
              lw = ser[:line_width] ? (ser[:line_width] * 12_700).to_i : nil
              xml.tag("a:ln", { w: lw, cap: ser[:line_cap] }) do
                xml.empty_tag("a:noFill") if ser[:no_line]
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, ser[:line_color]) } if ser[:line_color]
                xml.empty_tag("a:prstDash", { val: ser[:line_dash] }) if ser[:line_dash]
                write_line_join_xml(xml, ser[:line_join], ser[:line_miter_limit])
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped source) -> void
        def write_marker_xml(xml, source)
          has_sp = source[:marker_fill] || source[:marker_no_fill] || source[:marker_line_color] || source[:marker_line_dash] || source[:marker_no_line]
          return unless source[:marker_symbol] || source[:marker_size] || has_sp

          xml.tag("c:marker") do
            xml.empty_tag("c:symbol", { val: source[:marker_symbol] }) if source[:marker_symbol]
            xml.empty_tag("c:size", { val: source[:marker_size] }) if source[:marker_size]
            if has_sp
              xml.tag("c:spPr") do
                xml.empty_tag("a:noFill") if source[:marker_no_fill]
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, source[:marker_fill]) } if source[:marker_fill]
                if source[:marker_no_line]
                  xml.tag("a:ln") { xml.empty_tag("a:noFill") }
                elsif source[:marker_line_color] || source[:marker_line_dash]
                  lw = source[:marker_line_width] ? (source[:marker_line_width] * 12_700).to_i : nil
                  xml.tag("a:ln", { w: lw }) do
                    xml.tag("a:solidFill") { write_drawing_color_xml(xml, source[:marker_line_color]) } if source[:marker_line_color]
                    xml.empty_tag("a:prstDash", { val: source[:marker_line_dash] }) if source[:marker_line_dash]
                  end
                end
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped data_point) -> void
        def write_data_point_xml(xml, data_point)
          xml.tag("c:dPt") do
            xml.empty_tag("c:idx", { val: data_point[:idx] || "" })
            xml.empty_tag("c:invertIfNegative", { val: data_point[:invert_if_negative] ? 1 : 0 }) unless data_point[:invert_if_negative].nil?
            write_marker_xml(xml, data_point)
            xml.empty_tag("c:bubble3D", { val: data_point[:bubble_3d] ? 1 : 0 }) unless data_point[:bubble_3d].nil?
            xml.empty_tag("c:explosion", { val: data_point[:explosion] }) if data_point[:explosion]
            has_ln = data_point[:no_line] || data_point[:line_color] || data_point[:line_width] || data_point[:line_dash]
            if data_point[:fill_color] || data_point[:no_fill] || has_ln
              xml.tag("c:spPr") do
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, data_point[:fill_color]) } if data_point[:fill_color]
                xml.empty_tag("a:noFill") if data_point[:no_fill]
                if data_point[:no_line]
                  xml.tag("a:ln") { xml.empty_tag("a:noFill") }
                elsif data_point[:line_color] || data_point[:line_width] || data_point[:line_dash]
                  write_ln_xml(xml, line_color: data_point[:line_color], line_width: data_point[:line_width], line_dash: data_point[:line_dash])
                end
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped data_labels) -> void
        def write_data_labels_xml(xml, data_labels)
          xml.tag("c:dLbls") do
            data_labels[:labels]&.each do |lbl|
              xml.tag("c:dLbl") do
                xml.empty_tag("c:idx", { val: lbl[:idx] || "" })
                if lbl[:delete]
                  xml.empty_tag("c:delete", { val: 1 })
                else
                  if (lbl_layout = lbl[:layout])
                    ml_entries = [["c:x", lbl_layout[:x]], ["c:y", lbl_layout[:y]]].select { |_, val| val }
                    if ml_entries.empty?
                      xml.empty_tag("c:layout")
                    else
                      xml.tag("c:layout") do
                        xml.tag("c:manualLayout") do
                          ml_entries.each { |tag_name, val| xml.empty_tag(tag_name, { val: val }) }
                        end
                      end
                    end
                  end
                  if lbl[:text]
                    xml.tag("c:tx") do
                      xml.tag("c:rich") do
                        xml.empty_tag("a:bodyPr")
                        xml.empty_tag("a:lstStyle")
                        xml.tag("a:p") { xml.tag("a:r") { xml.text_tag("a:t", lbl[:text]) } }
                      end
                    end
                  end
                  if (lnf = lbl[:num_fmt])
                    xml.empty_tag("c:numFmt", { formatCode: lnf[:format_code] || "", sourceLinked: (1 if lnf[:source_linked]) })
                  end
                  write_simple_sp_pr_from_hash(xml, lbl)
                  write_axis_txpr(xml, nil, lbl[:font]) if lbl[:font]
                  write_data_label_flags_xml(xml, lbl)
                end
              end
            end
            if (nf = data_labels[:num_fmt])
              xml.empty_tag("c:numFmt", { formatCode: nf[:format_code] || "", sourceLinked: (1 if nf[:source_linked]) })
            end
            write_simple_sp_pr_from_hash(xml, data_labels)
            write_axis_txpr(xml, nil, data_labels[:font]) if data_labels[:font]
            write_data_label_flags_xml(xml, data_labels)
            xml.empty_tag("c:showLeaderLines", { val: data_labels[:show_leader_lines] ? 1 : 0 }) unless data_labels[:show_leader_lines].nil?
            write_line_sp_tag(xml, "c:leaderLines", data_labels[:leader_lines]) if data_labels[:leader_lines]
          end
        end

        # : (XmlBuilder xml, untyped lbl) -> void
        def write_data_label_flags_xml(xml, lbl)
          xml.empty_tag("c:dLblPos", { val: lbl[:position] }) if lbl[:position]
          xml.empty_tag("c:showLegendKey", { val: lbl[:show_legend_key] ? 1 : 0 }) unless lbl[:show_legend_key].nil?
          xml.empty_tag("c:showVal", { val: lbl[:show_val] ? 1 : 0 }) unless lbl[:show_val].nil?
          xml.empty_tag("c:showCatName", { val: lbl[:show_cat_name] ? 1 : 0 }) unless lbl[:show_cat_name].nil?
          xml.empty_tag("c:showSerName", { val: lbl[:show_ser_name] ? 1 : 0 }) unless lbl[:show_ser_name].nil?
          xml.empty_tag("c:showPercent", { val: lbl[:show_percent] ? 1 : 0 }) unless lbl[:show_percent].nil?
          xml.empty_tag("c:showBubbleSize", { val: lbl[:show_bubble_size] ? 1 : 0 }) unless lbl[:show_bubble_size].nil?
          xml.text_tag("c:separator", lbl[:separator]) if lbl[:separator]
        end

        # : (XmlBuilder xml, untyped trendline) -> void
        def write_trendline_xml(xml, trendline)
          xml.tag("c:trendline") do
            xml.text_tag("c:name", trendline[:name]) if trendline[:name]
            if trendline[:line_color] || trendline[:line_width] || trendline[:line_dash]
              xml.tag("c:spPr") do
                write_ln_xml(xml, line_color: trendline[:line_color], line_width: trendline[:line_width], line_dash: trendline[:line_dash])
              end
            end
            xml.empty_tag("c:trendlineType", { val: trendline[:type] || "linear" })
            xml.empty_tag("c:order", { val: trendline[:order] }) if trendline[:order]
            xml.empty_tag("c:period", { val: trendline[:period] }) if trendline[:period]
            xml.empty_tag("c:forward", { val: trendline[:forward] }) if trendline[:forward]
            xml.empty_tag("c:backward", { val: trendline[:backward] }) if trendline[:backward]
            xml.empty_tag("c:intercept", { val: trendline[:intercept] }) if trendline[:intercept]
            xml.empty_tag("c:dispRSqr", { val: trendline[:disp_r_sqr] ? 1 : 0 }) unless trendline[:disp_r_sqr].nil?
            xml.empty_tag("c:dispEq", { val: trendline[:disp_eq] ? 1 : 0 }) unless trendline[:disp_eq].nil?
            if (lbl = trendline[:label])
              xml.tag("c:trendlineLbl") do
                write_title_layout_xml(xml, lbl[:layout])
                if lbl[:text]
                  xml.tag("c:tx") do
                    xml.tag("c:rich") do
                      xml.empty_tag("a:bodyPr")
                      xml.empty_tag("a:lstStyle")
                      xml.tag("a:p") { xml.tag("a:r") { xml.text_tag("a:t", lbl[:text]) } }
                    end
                  end
                end
                if (nf = lbl[:num_fmt])
                  nf_src = nf.is_a?(Hash) ? nf : { format_code: nf }
                  nf_linked = nf_src[:source_linked] ? 1 : 0
                  xml.empty_tag("c:numFmt", { formatCode: nf_src[:format_code] || "", sourceLinked: nf_linked })
                end
                write_simple_sp_pr_from_hash(xml, lbl)
                write_axis_txpr(xml, nil, lbl[:font]) if lbl[:font]
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped err_bars) -> void
        def write_error_bars_xml(xml, err_bars)
          xml.tag("c:errBars") do
            xml.empty_tag("c:errDir", { val: err_bars[:direction] }) if err_bars[:direction]
            xml.empty_tag("c:errBarType", { val: err_bars[:bar_type] || "both" })
            xml.empty_tag("c:errValType", { val: err_bars[:val_type] || "fixedVal" })
            xml.empty_tag("c:noEndCap", { val: err_bars[:no_end_cap] ? 1 : 0 }) unless err_bars[:no_end_cap].nil?
            xml.tag("c:plus") { xml.tag("c:numRef") { xml.text_tag("c:f", err_bars[:plus]) } } if err_bars[:plus]
            xml.tag("c:minus") { xml.tag("c:numRef") { xml.text_tag("c:f", err_bars[:minus]) } } if err_bars[:minus]
            xml.empty_tag("c:val", { val: err_bars[:val] }) if err_bars[:val]
            write_simple_sp_pr_from_hash(xml, err_bars)
          end
        end

        # : (XmlBuilder xml, untyped band_fmts) -> void
        def write_band_fmts_xml(xml, band_fmts)
          xml.tag("c:bandFmts") do
            band_fmts.each do |bf|
              xml.tag("c:bandFmt") do
                xml.empty_tag("c:idx", { val: bf[:idx] || "" })
                write_simple_sp_pr_from_hash(xml, bf)
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped udb) -> void
        def write_up_down_bars_xml(xml, udb)
          xml.tag("c:upDownBars") do
            xml.empty_tag("c:gapWidth", { val: udb[:gap_width] }) if udb.is_a?(Hash) && udb[:gap_width]
            %i[up_bars down_bars].each do |bar_key|
              tag = bar_key == :up_bars ? "c:upBars" : "c:downBars"
              bar = udb.is_a?(Hash) ? udb[bar_key] : nil
              if bar && (bar[:fill_color] || bar[:no_fill] || bar[:line_color] || bar[:line_width] || bar[:line_dash])
                xml.tag(tag) { write_simple_sp_pr_from_hash(xml, bar) }
              else
                xml.empty_tag(tag)
              end
            end
          end
        end

        # : (XmlBuilder xml, untyped chart, ::Array[untyped] groups, bool has_secondary) -> void
        def write_chart_axes_xml(xml, chart, groups, has_secondary)
          cat_ax_tag = chart[:cat_axis_type] == :date ? "cat_date" : "cat_ax"
          el_name = cat_ax_tag == "cat_date" ? "c:dateAx" : "c:catAx"
          xml.tag(el_name) do
            xml.empty_tag("c:axId", { val: 1 })
            xml.tag("c:scaling") do
              xml.empty_tag("c:logBase", { val: chart[:cat_axis_log_base] }) if chart[:cat_axis_log_base]
              xml.empty_tag("c:orientation", { val: chart[:cat_axis_orientation] || "minMax" })
              xml.empty_tag("c:max", { val: chart[:cat_axis_scaling_max] }) if chart[:cat_axis_scaling_max]
              xml.empty_tag("c:min", { val: chart[:cat_axis_scaling_min] }) if chart[:cat_axis_scaling_min]
            end
            xml.empty_tag("c:delete", { val: chart[:cat_axis_delete] ? 1 : 0 })
            xml.empty_tag("c:axPos", { val: chart[:cat_axis_pos] || "b" })
            write_gridlines_xml(xml, "majorGridlines", chart[:cat_axis_major_gridlines]) if chart[:cat_axis_major_gridlines]
            write_gridlines_xml(xml, "minorGridlines", chart[:cat_axis_minor_gridlines]) if chart[:cat_axis_minor_gridlines]
            if chart[:cat_axis_title]
              cat_title_spec = merge_flat_title_styling(chart[:cat_axis_title], chart, :cat_axis_title)
              write_chart_title_xml(xml, cat_title_spec)
            end
            if (cnf = chart[:cat_axis_num_fmt])
              xml.empty_tag("c:numFmt", { formatCode: cnf[:format_code] || "", sourceLinked: cnf[:source_linked] ? 1 : 0 })
            end
            xml.empty_tag("c:majorTickMark", { val: chart[:cat_axis_major_tick_mark] }) if chart[:cat_axis_major_tick_mark]
            xml.empty_tag("c:minorTickMark", { val: chart[:cat_axis_minor_tick_mark] }) if chart[:cat_axis_minor_tick_mark]
            xml.empty_tag("c:tickLblPos", { val: chart[:cat_axis_tick_lbl_pos] }) if chart[:cat_axis_tick_lbl_pos]
            if chart[:cat_axis_line_color] || chart[:cat_axis_fill] || chart[:cat_axis_no_fill] || chart[:cat_axis_line_dash]
              xml.tag("c:spPr") do
                xml.empty_tag("a:noFill") if chart[:cat_axis_no_fill]
                xml.tag("a:solidFill") { write_drawing_color_xml(xml, chart[:cat_axis_fill]) } if chart[:cat_axis_fill]
                if chart[:cat_axis_line_color] || chart[:cat_axis_line_dash]
                  xml.tag("a:ln", { w: chart[:cat_axis_line_width] }) do
                    xml.tag("a:solidFill") { write_drawing_color_xml(xml, chart[:cat_axis_line_color]) } if chart[:cat_axis_line_color]
                    xml.empty_tag("a:prstDash", { val: chart[:cat_axis_line_dash] }) if chart[:cat_axis_line_dash]
                  end
                end
              end
            end
            write_axis_txpr(xml, chart[:cat_axis_label_rotation], chart[:cat_axis_font])
            xml.empty_tag("c:crossAx", { val: 2 })
            xml.empty_tag("c:crosses", { val: chart[:cat_axis_crosses] }) if chart[:cat_axis_crosses]
            xml.empty_tag("c:crossesAt", { val: chart[:cat_axis_crosses_at] }) if !chart[:cat_axis_crosses] && chart[:cat_axis_crosses_at]
            xml.empty_tag("c:auto", { val: chart[:cat_axis_auto] ? 1 : 0 }) unless chart[:cat_axis_auto].nil?
            xml.empty_tag("c:lblAlgn", { val: chart[:cat_axis_lbl_algn] }) if chart[:cat_axis_lbl_algn]
            xml.empty_tag("c:lblOffset", { val: chart[:cat_axis_lbl_offset] }) if chart[:cat_axis_lbl_offset]
            xml.empty_tag("c:tickLblSkip", { val: chart[:cat_axis_tick_lbl_skip] }) if chart[:cat_axis_tick_lbl_skip]
            xml.empty_tag("c:tickMarkSkip", { val: chart[:cat_axis_tick_mark_skip] }) if chart[:cat_axis_tick_mark_skip]
            xml.empty_tag("c:noMultiLvlLbl", { val: chart[:cat_axis_no_multi_lvl_lbl] ? 1 : 0 }) unless chart[:cat_axis_no_multi_lvl_lbl].nil?
            if el_name == "c:dateAx"
              xml.empty_tag("c:baseTimeUnit", { val: chart[:cat_axis_base_time_unit] }) if chart[:cat_axis_base_time_unit]
              xml.empty_tag("c:majorUnit", { val: chart[:cat_axis_major_unit] }) if chart[:cat_axis_major_unit]
              xml.empty_tag("c:majorTimeUnit", { val: chart[:cat_axis_major_time_unit] }) if chart[:cat_axis_major_time_unit]
              xml.empty_tag("c:minorUnit", { val: chart[:cat_axis_minor_unit] }) if chart[:cat_axis_minor_unit]
              xml.empty_tag("c:minorTimeUnit", { val: chart[:cat_axis_minor_time_unit] }) if chart[:cat_axis_minor_time_unit]
            end
          end

          write_val_ax_xml(xml, ax_id: 2, prefix: :val_axis, chart: chart, default_pos: "l")
          write_val_ax_xml(xml, ax_id: 3, prefix: :sec_val_axis, chart: chart, default_pos: "r", default_crosses: "max") if has_secondary

          return unless groups.any? { |g| THREE_AXIS_CHARTS.include?(g[:type]) }

          xml.tag("c:serAx") do
            xml.empty_tag("c:axId", { val: 3 })
            xml.tag("c:scaling") { xml.empty_tag("c:orientation", { val: "minMax" }) }
            xml.empty_tag("c:delete", { val: 0 })
            xml.empty_tag("c:axPos", { val: "b" })
            xml.empty_tag("c:crossAx", { val: 2 })
          end
        end

        # : (XmlBuilder xml, untyped data_table) -> void
        def write_chart_data_table_xml(xml, data_table)
          xml.tag("c:dTable") do
            xml.empty_tag("c:showHorzBorder", { val: data_table[:show_horz_border] ? 1 : 0 }) unless data_table[:show_horz_border].nil?
            xml.empty_tag("c:showVertBorder", { val: data_table[:show_vert_border] ? 1 : 0 }) unless data_table[:show_vert_border].nil?
            xml.empty_tag("c:showOutline", { val: data_table[:show_outline] ? 1 : 0 }) unless data_table[:show_outline].nil?
            xml.empty_tag("c:showKeys", { val: data_table[:show_keys] ? 1 : 0 }) unless data_table[:show_keys].nil?
            write_simple_sp_pr_from_hash(xml, data_table)
            write_axis_txpr(xml, nil, data_table[:font]) if data_table[:font]
          end
        end

        # : (XmlBuilder xml, untyped legend) -> void
        def write_chart_legend_xml(xml, legend)
          legend_pos = legend&.[](:position) || "r"
          legend_overlay = legend&.[](:overlay)
          xml.tag("c:legend") do
            xml.empty_tag("c:legendPos", { val: legend_pos })
            legend&.[](:entries)&.each do |entry|
              xml.tag("c:legendEntry") do
                xml.empty_tag("c:idx", { val: entry[:idx] || "" })
                if !entry[:delete].nil?
                  xml.empty_tag("c:delete", { val: entry[:delete] ? 1 : 0 })
                elsif entry[:font]
                  write_axis_txpr(xml, nil, entry[:font])
                end
              end
            end
            write_plot_area_layout_xml(xml, legend&.[](:layout), emit_empty: false)
            xml.empty_tag("c:overlay", { val: legend_overlay ? 1 : 0 }) unless legend_overlay.nil?
            write_simple_sp_pr_from_hash(xml, legend) if legend
            write_axis_txpr(xml, nil, legend[:font]) if legend&.[](:font)
          end
        end

        # : (XmlBuilder xml, untyped print_settings) -> void
        def write_chart_print_settings_xml(xml, print_settings)
          xml.tag("c:printSettings") do
            if (hf = print_settings[:header_footer])
              xml.tag("c:headerFooter") do
                xml.text_tag("c:oddHeader", hf[:odd_header]) if hf[:odd_header]
                xml.text_tag("c:oddFooter", hf[:odd_footer]) if hf[:odd_footer]
                xml.text_tag("c:evenHeader", hf[:even_header]) if hf[:even_header]
                xml.text_tag("c:evenFooter", hf[:even_footer]) if hf[:even_footer]
                xml.text_tag("c:firstHeader", hf[:first_header]) if hf[:first_header]
                xml.text_tag("c:firstFooter", hf[:first_footer]) if hf[:first_footer]
              end
            end
            if (pm = print_settings[:page_margins])
              pm_attrs = %w[b l r t header footer].each_with_object({}) do |a, h|
                val = pm[a.to_sym]
                h[a] = val if val
              end
              xml.empty_tag("c:pageMargins", pm_attrs)
            end
            if (psu = print_settings[:page_setup])
              xml.empty_tag(
                "c:pageSetup",
                {
                  paperSize: psu[:paper_size],
                  firstPageNumber: psu[:first_page_number],
                  orientation: psu[:orientation],
                  horizontalDpi: psu[:horizontal_dpi],
                  verticalDpi: psu[:vertical_dpi],
                  copies: psu[:copies]
                }
              )
            end
          end
        end

        # : (XmlBuilder xml, untyped hash) -> void
        def write_simple_sp_pr_from_hash(xml, hash)
          return unless hash[:fill_color] || hash[:no_fill] || hash[:line_color] || hash[:line_width] || hash[:line_dash]

          write_simple_sp_pr(
            xml,
            fill_color: hash[:fill_color],
            no_fill: hash[:no_fill],
            line_color: hash[:line_color],
            line_width: hash[:line_width],
            line_dash: hash[:line_dash]
          )
        end

        # : (XmlBuilder xml, ?fill_color: untyped?, ?no_fill: untyped?, ?line_color: untyped?, ?line_width: untyped?, ?line_dash: untyped?, ?emu: bool) -> void
        def write_simple_sp_pr(xml, fill_color: nil, no_fill: false, line_color: nil, line_width: nil, line_dash: nil, emu: true)
          xml.tag("c:spPr") do
            xml.tag("a:solidFill") { write_drawing_color_xml(xml, fill_color) } if fill_color
            xml.empty_tag("a:noFill") if no_fill
            write_ln_xml(xml, line_color: line_color, line_width: line_width, line_dash: line_dash, emu: emu)
          end
        end

        # : (XmlBuilder xml, ?line_color: untyped?, ?line_width: untyped?, ?line_dash: untyped?, ?emu: bool) -> void
        def write_ln_xml(xml, line_color: nil, line_width: nil, line_dash: nil, emu: true)
          return unless line_color || line_width || line_dash

          w = if line_width
                emu ? (line_width * 12_700).to_i : line_width.to_i
              end
          xml.tag("a:ln", { w: w }) do
            xml.tag("a:solidFill") { write_drawing_color_xml(xml, line_color) } if line_color
            xml.empty_tag("a:prstDash", { val: line_dash }) if line_dash
          end
        end

        # : (XmlBuilder xml, ::String tag, untyped spec) -> void
        def write_line_sp_tag(xml, tag, spec)
          if spec.is_a?(Hash) && (spec[:line_color] || spec[:line_width] || spec[:line_dash])
            xml.tag(tag) do
              xml.tag("c:spPr") do
                write_ln_xml(xml, line_color: spec[:line_color], line_width: spec[:line_width], line_dash: spec[:line_dash])
              end
            end
          else
            xml.empty_tag(tag)
          end
        end
      end
    end
  end
end
