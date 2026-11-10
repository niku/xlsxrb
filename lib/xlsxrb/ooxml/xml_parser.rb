# frozen_string_literal: true

# rbs_inline: enabled

require "rexml/parsers/sax2parser"
require "rexml/sax2listener"

# Ensure REXML constants and security limits are accessible across Ractors
begin
  if defined?(Ractor)
    Ractor.make_shareable(REXML::Parsers::BaseParser::DEFAULT_ENTITIES) if defined?(REXML::Parsers::BaseParser::DEFAULT_ENTITIES) && !Ractor.shareable?(REXML::Parsers::BaseParser::DEFAULT_ENTITIES)
    if defined?(REXML::Source) && REXML::Source.const_defined?(:Private)
      priv = REXML::Source.const_get(:Private)
      if priv.const_defined?(:PRE_DEFINED_TERM_PATTERNS)
        terms = priv.const_get(:PRE_DEFINED_TERM_PATTERNS)
        Ractor.make_shareable(terms) unless Ractor.shareable?(terms)
      end
    end
    if defined?(REXML::Security)
      # REXML keeps these limits in class variables, which non-main Ractors cannot read.
      # Move them to module instance variables, carrying over the values the application
      # has already set (e.g. 0 to disable entity expansion).
      current_limit = REXML::Security.entity_expansion_limit
      current_text_limit = REXML::Security.entity_expansion_text_limit
      prev_verbose = $VERBOSE
      $VERBOSE = nil
      begin
        REXML::Security.instance_eval do
          @entity_expansion_limit = current_limit
          @entity_expansion_text_limit = current_text_limit

          def self.entity_expansion_limit
            @entity_expansion_limit
          end

          def self.entity_expansion_limit=(val)
            @entity_expansion_limit = val
          end

          def self.entity_expansion_text_limit
            @entity_expansion_text_limit
          end

          def self.entity_expansion_text_limit=(val)
            @entity_expansion_text_limit = val
          end
        end
        Ractor.make_shareable(REXML::Security)
      ensure
        $VERBOSE = prev_verbose
      end
    end
  end
rescue StandardError
  # Best-effort Ractor isolation preparation for standard library REXML
end

module Xlsxrb
  module Ooxml
    # Thin wrapper around REXML SAX2 parser for streaming XML processing.
    # Unknown elements are captured as unmapped data hashes.
    module XmlParser
      # Parses XML string using SAX2, calling the given listener.
      def self.parse(xml_string, listener)
        parser = sax2_parser(xml_string)
        parser.listen(listener)
        parser.parse
      end

      # Returns a SAX2 parser whose listeners receive attribute values with XML entities
      # decoded. REXML's SAX2 parser decodes text but passes attribute values as written, so
      # formatCode="&quot;$&quot;0" would otherwise arrive with the &quot; references.
      #: (String xml_string) -> REXML::Parsers::SAX2Parser
      def self.sax2_parser(xml_string)
        DecodingSAX2Parser.new(xml_string)
      end

      # SAX2 parser that wraps each listener in an AttributeDecoder.
      class DecodingSAX2Parser < REXML::Parsers::SAX2Parser
        #: (*untyped args) ?{ (*untyped) -> untyped } -> untyped
        def listen(*args, &)
          return super if args.size != 1 || args.first.is_a?(Symbol) || args.first.is_a?(Array) || block_given?

          super(AttributeDecoder.new(args.first))
        end
      end

      # Forwards SAX2 events to a listener, decoding entity references in attribute values.
      class AttributeDecoder
        #: (untyped listener) -> void
        def initialize(listener)
          @listener = listener
        end

        # Plain methods rather than define_method blocks, which Ractors cannot share.
        #: (*untyped args) -> untyped
        def start_document(*args) = @listener.start_document(*args)
        #: (*untyped args) -> untyped
        def end_document(*args) = @listener.end_document(*args)
        #: (*untyped args) -> untyped
        def start_prefix_mapping(*args) = @listener.start_prefix_mapping(*args)
        #: (*untyped args) -> untyped
        def end_prefix_mapping(*args) = @listener.end_prefix_mapping(*args)
        #: (*untyped args) -> untyped
        def end_element(*args) = @listener.end_element(*args)
        #: (*untyped args) -> untyped
        def characters(*args) = @listener.characters(*args)
        #: (*untyped args) -> untyped
        def processing_instruction(*args) = @listener.processing_instruction(*args)
        #: (*untyped args) -> untyped
        def doctype(*args) = @listener.doctype(*args)
        #: (*untyped args) -> untyped
        def attlistdecl(*args) = @listener.attlistdecl(*args)
        #: (*untyped args) -> untyped
        def elementdecl(*args) = @listener.elementdecl(*args)
        #: (*untyped args) -> untyped
        def entitydecl(*args) = @listener.entitydecl(*args)
        #: (*untyped args) -> untyped
        def notationdecl(*args) = @listener.notationdecl(*args)
        #: (*untyped args) -> untyped
        def cdata(*args) = @listener.cdata(*args)
        #: (*untyped args) -> untyped
        def xmldecl(*args) = @listener.xmldecl(*args)
        #: (*untyped args) -> untyped
        def comment(*args) = @listener.comment(*args)
        #: (*untyped args) -> untyped
        def progress(*args) = @listener.progress(*args)

        #: (String? uri, String local_name, String qname, Hash[String, String] attributes) -> untyped
        def start_element(uri, local_name, qname, attributes)
          if attributes.each_value.any? { |value| value.include?("&") }
            attributes = attributes.transform_values { |value| value.include?("&") ? REXML::Text.unnormalize(value) : value }
          end
          @listener.start_element(uri, local_name, qname, attributes)
        end
      end

      # Base listener with unmapped-data collection support.
      class BaseListener
        include REXML::SAX2Listener

        attr_reader :unmapped_data

        def initialize
          @unmapped_data = []
          @unmapped_stack = []
          @capturing_unmapped = false
          @current_unmapped = nil
        end

        # Override in subclass: return true if the tag is recognized.
        def recognized_tag?(_uri, _localname, _qname)
          false
        end

        def start_element(uri, localname, qname, attrs)
          if @capturing_unmapped
            child = { tag: localname, attrs: attrs.dup, children: [], text: nil }
            @unmapped_stack.last[:children] << child
            @unmapped_stack.push(child)
          elsif !recognized_tag?(uri, localname, qname)
            @capturing_unmapped = true
            @current_unmapped = { tag: localname, attrs: attrs.dup, children: [], text: nil }
            @unmapped_stack.push(@current_unmapped)
          end
        end

        def end_element(_uri, _localname, _qname)
          return unless @capturing_unmapped

          @unmapped_stack.pop
          return unless @unmapped_stack.empty?

          @capturing_unmapped = false
          @unmapped_data << @current_unmapped
          @current_unmapped = nil
        end

        def characters(text)
          return unless @capturing_unmapped && !@unmapped_stack.empty?

          current = @unmapped_stack.last
          current[:text] = (current[:text] || "") + text
        end
      end
    end
  end
end
