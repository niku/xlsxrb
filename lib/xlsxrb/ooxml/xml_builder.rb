# frozen_string_literal: true

module Xlsxrb
  module Ooxml
    # Streams well-formed XML to a writable IO without building a DOM.
    class XmlBuilder
      XML_HEADER = %(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n)

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

      def declaration
        @io << XML_HEADER
        self
      end

      # Opens a tag, yields for children, then closes the tag.
      def tag(name, attrs = {}, &block)
        if block
          open_tag(name, attrs)
          yield self
          close_tag(name)
        else
          empty_tag(name, attrs)
        end
        self
      end

      def open_tag(name, attrs = {})
        @io << "<#{name}"
        write_attrs(attrs)
        @io << ">"
        self
      end

      def close_tag(name)
        @io << "</#{name}>"
        self
      end

      def empty_tag(name, attrs = {})
        @io << "<#{name}"
        write_attrs(attrs)
        @io << "/>"
        self
      end

      def text(content)
        @io << escape(content.to_s)
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

      private

      def write_attrs(attrs)
        attrs.each do |k, v|
          next if v.nil?

          @io << %( #{k}="#{escape(v.to_s)}")
        end
      end

      def escape(str)
        str.gsub(ESCAPE_RE, ESCAPE_MAP)
      end
    end
  end
end
