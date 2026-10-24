# frozen_string_literal: true

# rbs_inline: enabled

require_relative "xml_parser"
require_relative "xml_builder"

module Xlsxrb
  module Ooxml
    # SAX-based parser for xl/sharedStrings.xml.
    # Returns an Array of strings (index = SST index).
    class SharedStringsParser
      TAG_DELIM_BYTES = [62, 32, 47, 9, 10, 13].freeze # '>', ' ', '/', '\t', '\n', '\r'
      private_constant :TAG_DELIM_BYTES

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
        sst_term = xml.index("sst")
        return unless sst_term

        sst_start = xml.rindex("<", sst_term)
        return unless sst_start

        prefix = xml.byteslice(sst_start + 1, sst_term - (sst_start + 1))
        return if !prefix.empty? && !prefix.end_with?(":")

        sst_open_end = xml.index(">", sst_start)
        return unless sst_open_end
        return if xml.getbyte(sst_open_end - 1) == 47 # self-closing <sst/>

        sst_end = xml.index("</#{prefix}sst>", sst_open_end)
        sst_end ||= xml.bytesize

        si_start_pattern = "<#{prefix}si"
        si_start_len = si_start_pattern.bytesize
        si_end_tag = "</#{prefix}si>"
        si_end_len = si_end_tag.bytesize

        t_start_pattern = "<#{prefix}t"
        t_start_len = t_start_pattern.bytesize
        t_end_tag = "</#{prefix}t>"
        t_end_len = t_end_tag.bytesize

        rph_start_pattern = "<#{prefix}rPh"
        rph_end_tag = "</#{prefix}rPh>"
        rph_end_len = rph_end_tag.bytesize
        has_rph = xml.include?(rph_start_pattern)

        pos = sst_open_end + 1
        while pos < sst_end
          si_start = xml.index(si_start_pattern, pos)
          break unless si_start && si_start < sst_end

          si_open_end = xml.index(">", si_start + si_start_len)
          break unless si_open_end

          # Self-closing <si/>
          if xml.getbyte(si_open_end - 1) == 47
            yield ""
            pos = si_open_end + 1
            next
          end

          si_end = xml.index(si_end_tag, si_open_end + 1)
          break unless si_end

          # Fast path: check for simple single <t>...</t> inside <si>
          first_t = xml.index(t_start_pattern, si_open_end + 1)
          if first_t && first_t < si_end
            tb = xml.getbyte(first_t + t_start_len)
            if TAG_DELIM_BYTES.include?(tb)
              t_open_end = xml.index(">", first_t + t_start_len)
              if t_open_end && t_open_end < si_end
                if xml.getbyte(t_open_end - 1) == 47 # <t/>
                  next_t = xml.index(t_start_pattern, t_open_end + 1)
                  str = if next_t && next_t < si_end
                          extract_multi_t(xml, si_open_end + 1, si_end, t_start_pattern, t_start_len, t_end_tag, t_end_len, rph_start_pattern, rph_end_tag, rph_end_len, has_rph: has_rph)
                        else
                          ""
                        end
                else
                  t_end = xml.index(t_end_tag, t_open_end + 1)
                  if t_end && t_end <= si_end
                    next_t = xml.index(t_start_pattern, t_end + t_end_len)
                    if next_t && next_t < si_end
                      # Multiple <t> tags (rich text runs)
                      str = extract_multi_t(xml, si_open_end + 1, si_end, t_start_pattern, t_start_len, t_end_tag, t_end_len, rph_start_pattern, rph_end_tag, rph_end_len, has_rph: has_rph)
                    else
                      raw_str = xml.byteslice(t_open_end + 1, t_end - t_open_end - 1).force_encoding("UTF-8")
                      str = XmlBuilder.unescape(raw_str)
                    end
                  else
                    str = ""
                  end
                end
              else
                str = ""
              end
            else
              str = extract_multi_t(xml, si_open_end + 1, si_end, t_start_pattern, t_start_len, t_end_tag, t_end_len, rph_start_pattern, rph_end_tag, rph_end_len, has_rph: has_rph)
            end
          else
            str = ""
          end

          yield normalize_newlines(str).freeze
          pos = si_end + si_end_len
        end
      end

      private_class_method :fast_scan

      def self.extract_multi_t(xml, from, to, t_start_pattern, t_start_len, t_end_tag, t_end_len, rph_start_pattern, rph_end_tag, rph_end_len, has_rph: false)
        buf = +""
        pos = from
        while pos < to
          t_start = xml.index(t_start_pattern, pos)
          break unless t_start && t_start < to

          # Ensure the tag is actually <t>, <t ...>, or <t/> and not <theme>, <table>, etc.
          b = xml.getbyte(t_start + t_start_len)
          # rubocop:disable Style/MultipleComparison -- Zero-allocation hot path check
          unless b == 62 || b == 32 || b == 47 || b == 9 || b == 10 || b == 13
            pos = t_start + t_start_len
            next
          end
          # rubocop:enable Style/MultipleComparison

          if has_rph
            # Phonetic runs (<rPh>) hold furigana readings and are not part of the string
            rph_start = xml.index(rph_start_pattern, pos)
            if rph_start && rph_start < t_start
              rph_end = xml.index(rph_end_tag, rph_start)
              break unless rph_end && rph_end < to

              pos = rph_end + rph_end_len
              next
            end
          end

          t_open_end = xml.index(">", t_start + t_start_len)
          break unless t_open_end

          if xml.getbyte(t_open_end - 1) == 47
            pos = t_open_end + 1
            next
          end

          t_end = xml.index(t_end_tag, t_open_end + 1)
          break unless t_end && t_end <= to

          raw_chunk = xml.byteslice(t_open_end + 1, t_end - t_open_end - 1).force_encoding("UTF-8")
          buf << XmlBuilder.unescape(raw_chunk)
          pos = t_end + t_end_len
        end
        buf
      end

      private_class_method :extract_multi_t

      def self.normalize_newlines(str)
        str.include?("\r") ? str.gsub("\r\n", "\n").tr("\r", "\n") : str
      end

      private_class_method :normalize_newlines
    end
  end
end
