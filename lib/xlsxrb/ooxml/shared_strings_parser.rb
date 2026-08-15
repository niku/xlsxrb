# frozen_string_literal: true

# rbs_inline: enabled

require_relative "xml_parser"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/sharedStrings.xml.
    # Returns an Array of strings (index = SST index).
    class SharedStringsParser
      XML_ENTITIES = { "&amp;" => "&", "&lt;" => "<", "&gt;" => ">", "&quot;" => '"', "&apos;" => "'" }.freeze

      # Parses all shared strings and returns an Array of strings.
      def self.parse(xml_string, _part_name: "xl/sharedStrings.xml")
        return [] if xml_string.nil? || xml_string.empty?

        if xml_string.include?("<!DOCTYPE") || xml_string.include?("<!ENTITY")
          listener = Class.new { include REXML::SAX2Listener }.new
          XmlParser.parse(xml_string, listener)
        end

        strings = []
        fast_scan(xml_string) do |str|
          strings << str
        end
        strings
      end

      # Yields Event objects for each shared string.
      def self.each_event(xml_string, part_name: "xl/sharedStrings.xml", &block)
        return enum_for(:each_event, xml_string, part_name: part_name) unless block
        return if xml_string.nil? || xml_string.empty?

        index = 0
        fast_scan(xml_string) do |str|
          block.call(Event.new(
                       type: :sst_item,
                       args: [str],
                       source: { part: part_name, index: index }
                     ))
          index += 1
        end
      end

      def self.fast_scan(xml_src)
        xml = xml_src.b
        sst_start = xml.index("<sst")
        return unless sst_start

        sst_open_end = xml.index(">", sst_start)
        return unless sst_open_end
        return if xml.getbyte(sst_open_end - 1) == 47 # self-closing <sst/>

        sst_end = xml.index("</sst>", sst_open_end)
        sst_end ||= xml.bytesize

        pos = sst_open_end + 1
        while pos < sst_end
          si_start = xml.index("<si", pos)
          break unless si_start && si_start < sst_end

          si_open_end = xml.index(">", si_start + 3)
          break unless si_open_end

          # Self-closing <si/>
          if xml.getbyte(si_open_end - 1) == 47
            yield ""
            pos = si_open_end + 1
            next
          end

          si_end = xml.index("</si>", si_open_end + 1)
          break unless si_end

          # Fast path: check for simple single <t>...</t> inside <si>
          first_t = xml.index("<t", si_open_end + 1)
          if first_t && first_t < si_end
            t_open_end = xml.index(">", first_t + 2)
            if t_open_end && t_open_end < si_end
              if xml.getbyte(t_open_end - 1) == 47 # <t/>
                str = ""
              else
                t_end = xml.index("</t>", t_open_end + 1)
                if t_end && t_end <= si_end
                  next_t = xml.index("<t", t_end + 4)
                  if next_t && next_t < si_end
                    # Multiple <t> tags (rich text runs)
                    str = extract_multi_t(xml, si_open_end + 1, si_end)
                  else
                    raw_str = xml.byteslice(t_open_end + 1, t_end - t_open_end - 1).force_encoding("UTF-8")
                    str = raw_str.include?("&") ? raw_str.gsub(/&(?:amp|lt|gt|quot|apos);/, XML_ENTITIES) : raw_str
                  end
                else
                  str = ""
                end
              end
            else
              str = ""
            end
          else
            str = ""
          end

          yield str.freeze
          pos = si_end + 5
        end
      end

      private_class_method :fast_scan

      def self.extract_multi_t(xml, from, to)
        buf = +""
        pos = from
        while pos < to
          t_start = xml.index("<t", pos)
          break unless t_start && t_start < to

          t_open_end = xml.index(">", t_start + 2)
          break unless t_open_end

          if xml.getbyte(t_open_end - 1) == 47
            pos = t_open_end + 1
            next
          end

          t_end = xml.index("</t>", t_open_end + 1)
          break unless t_end && t_end <= to

          raw_chunk = xml.byteslice(t_open_end + 1, t_end - t_open_end - 1).force_encoding("UTF-8")
          raw_chunk = raw_chunk.gsub(/&(?:amp|lt|gt|quot|apos);/, XML_ENTITIES) if raw_chunk.include?("&")
          buf << raw_chunk
          pos = t_end + 4
        end
        buf
      end

      private_class_method :extract_multi_t
    end
  end
end
