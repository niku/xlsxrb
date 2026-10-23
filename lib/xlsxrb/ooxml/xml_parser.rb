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
      prev_verbose = $VERBOSE
      $VERBOSE = nil
      begin
        # rubocop:disable Style/Documentation
        module ::REXML
          module Security
            @entity_expansion_limit = 10_000
            @entity_expansion_text_limit = 10_240

            class << self
              attr_writer :entity_expansion_limit, :entity_expansion_text_limit

              def entity_expansion_limit
                @entity_expansion_limit || 10_000
              end

              def entity_expansion_text_limit
                @entity_expansion_text_limit || 10_240
              end
            end
          end
        end
        # rubocop:enable Style/Documentation
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
        parser = REXML::Parsers::SAX2Parser.new(xml_string)
        parser.listen(listener)
        parser.parse
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
