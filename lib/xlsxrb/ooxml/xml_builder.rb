# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Ooxml
    # Streams well-formed XML to a writable IO without building a DOM.
    class XmlBuilder
      XML_HEADER = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n)
      PLAIN_XML_HEADER = %(<?xml version="1.0" encoding="UTF-8"?>)

      # Builds an XML string with a new {XmlBuilder}.
      #
      # @param declaration [Symbol, false, nil] `:standalone`, `:plain`, or falsy for none.
      # @return [String]
      #: (?declaration: (:standalone | :plain | false)?) { (XmlBuilder) -> void } -> String
      def self.build(declaration: nil)
        io = +""
        builder = new(io)
        builder.declaration(standalone: declaration == :standalone) if declaration
        yield builder
        io
      end

      ESCAPE_MAP = {
        "&" => "&amp;",
        "<" => "&lt;",
        ">" => "&gt;",
        '"' => "&quot;",
        "'" => "&apos;"
      }.freeze

      ESCAPE_RE = /[&<>"']/

      def initialize(io)
        @io = io
      end

      #: (?standalone: bool) -> self
      def declaration(standalone: true)
        @io << (standalone ? XML_HEADER : PLAIN_XML_HEADER)
        self
      end

      EMPTY_HASH = {}.freeze

      # Opens a tag, yields for children, then closes the tag.
      def tag(name, attrs = EMPTY_HASH, &block)
        if block
          open_tag(name, attrs)
          yield self
          close_tag(name)
        else
          empty_tag(name, attrs)
        end
        self
      end

      def open_tag(name, attrs = EMPTY_HASH)
        @io << "<"
        @io << name
        write_attrs(attrs) unless attrs.empty?
        @io << ">"
        self
      end

      def close_tag(name)
        @io << "</"
        @io << name
        @io << ">"
        self
      end

      def empty_tag(name, attrs = EMPTY_HASH)
        @io << "<"
        @io << name
        write_attrs(attrs) unless attrs.empty?
        @io << "/>"
        self
      end

      def text(content)
        @io << escape(content.to_s)
        self
      end

      # Writes an element containing escaped text content.
      #: (String name, untyped content, ?Hash[untyped, untyped] attrs) -> self
      def text_tag(name, content, attrs = EMPTY_HASH)
        open_tag(name, attrs)
        text(content)
        close_tag(name)
      end

      # Writes ST_Xstring-encoded and XML-escaped text content.
      #: (untyped content) -> self
      def xstring_text(content)
        str = content.to_s
        str = str.dup.force_encoding(Encoding::UTF_8) if str.encoding.equal?(Encoding::BINARY)
        @io << escape(self.class.encode_xstring(str))
        self
      end

      # Write raw XML string (for unmapped_data restoration).
      def raw(xml_string)
        @io << xml_string
        self
      end

      # Serialize an unmapped_data hash back to XML.
      def write_unmapped(node)
        return unless node.is_a?(Hash) && node[:tag]

        tag_name = node[:tag]
        attrs = node[:attrs] || {}
        children = node[:children] || []
        text_content = node[:text]

        if children.empty? && (text_content.nil? || text_content.empty?)
          empty_tag(tag_name, attrs)
        else
          open_tag(tag_name, attrs)
          text(text_content) if text_content && !text_content.empty?
          children.each { |child| write_unmapped(child) }
          close_tag(tag_name)
        end
      end

      def to_s
        @io.is_a?(StringIO) ? @io.string : @io.to_s
      end

      UNESCAPE_MAP = {
        "&amp;" => "&",
        "&lt;" => "<",
        "&gt;" => ">",
        "&quot;" => '"',
        "&apos;" => "'"
      }.freeze

      UNESCAPE_RE = /&(?:amp|lt|gt|quot|apos);/

      INVALID_XML_CHARS_RE = /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/
      private_constant :INVALID_XML_CHARS_RE

      # Escapes XML special characters in a string, and removes characters that XML 1.0
      # cannot hold, such as U+0001 and U+FFFE: a part that contains one is not well-formed.
      # Binary strings are only escaped.
      #
      # @param value [Object]
      # @return [String]
      #: (untyped value) -> String
      def self.escape(value)
        str = value.to_s
        str = str.gsub(INVALID_XML_CHARS_RE, "") if !str.encoding.equal?(Encoding::BINARY) && str.match?(INVALID_XML_CHARS_RE)
        str.match?(ESCAPE_RE) ? str.gsub(ESCAPE_RE, ESCAPE_MAP) : str
      end

      XSTRING_ENCODE_RE = /_(?=x\h{4}_)|[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/
      private_constant :XSTRING_ENCODE_RE

      # Encodes cell text as ISO/IEC 29500-1 ST_Xstring, the inverse of the _xHHHH_ part
      # of unescape: an underscore that starts an _xHHHH_ sequence becomes _x005F_, and a
      # character that XML 1.0 cannot hold, such as U+0001 or U+FFFE, becomes _xHHHH_.
      # Excel and unescape read the original text back. The value must be UTF-8.
      #
      # @param value [String]
      # @return [String]
      #: (String value) -> String
      def self.encode_xstring(value)
        return value unless value.match?(XSTRING_ENCODE_RE)

        # The underscore is U+005F, so the same format gives _x005F_.
        value.gsub(XSTRING_ENCODE_RE) { |char| format("_x%04X_", char.ord) }
      end

      ENTITY_RE = /&(?:amp|lt|gt|quot|apos|#(\d+)|#x([0-9a-fA-F]+));/
      XSTRING_RE = /_x([0-9A-Fa-f]{4})_/
      private_constant :ENTITY_RE, :XSTRING_RE

      # Unescapes predefined and numeric/hex XML entities, as well as
      # ISO/IEC 29500-1 ST_Xstring (_xHHHH_) control character sequences.
      # Each entity is decoded once: "&amp;lt;" is the text "&lt;", not "<".
      #
      # @param value [Object]
      # @return [String]
      #: (untyped value) -> String
      def self.unescape(value)
        str = value.to_s
        return str unless str.include?("&") || str.include?("_x")

        # Predefined and numeric entities in one pass, so that the output of one is not read
        # as the start of another.
        res = str.gsub(ENTITY_RE) do |match|
          dec = Regexp.last_match(1)
          hex = Regexp.last_match(2)
          next UNESCAPE_MAP.fetch(match) unless dec || hex

          (dec ? Integer(dec) : Integer(hex, 16)).chr(Encoding::UTF_8)
        rescue RangeError
          match
        end
        res.gsub(XSTRING_RE) do |match|
          codepoint = Integer(Regexp.last_match(1), 16)
          codepoint.chr(Encoding::UTF_8)
        rescue RangeError
          match
        end
      end

      private

      def write_attrs(attrs)
        attrs.each do |k, v|
          next if v.nil?

          @io << " "
          @io << k.to_s
          @io << '="'
          @io << escape(v)
          @io << '"'
        end
      end

      def escape(value)
        self.class.escape(value)
      end
    end
  end
end
