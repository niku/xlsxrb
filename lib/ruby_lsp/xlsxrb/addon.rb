# frozen_string_literal: true

require "ruby_lsp/addon"
require_relative "completion_listener"
require_relative "../../xlsxrb/version"

module RubyLsp
  module Xlsxrb
    # Ruby LSP Add-on for xlsxrb.
    #
    # [Context & Lifecycle Note]
    # This add-on serves as a bridge/polyfill for current Ruby LSP environments.
    # While xlsxrb ships with complete RBS signatures (`sig/generated/`), Ruby LSP's
    # type inferrer does not yet perform automatic static type inference from method
    # block signatures to block parameters (e.g., `Xlsxrb.generate do |wb|`).
    #
    # This add-on enables immediate out-of-the-box autocompletion and rich Markdown
    # documentation across all public block arguments.
    #
    # As the Ruby tooling ecosystem evolves and Ruby LSP gains native block argument
    # type resolution from gem RBS signatures in the future, this add-on will become
    # redundant and can eventually be deprecated or removed.
    class Addon < ::RubyLsp::Addon
      def activate(global_state, _message_queue)
        @global_state = global_state
      end

      def deactivate; end

      def name
        "xlsxrb"
      end

      def version
        ::Xlsxrb::VERSION
      end

      def create_completion_listener(response_builder, node_context, dispatcher, _uri = nil)
        CompletionListener.new(response_builder, node_context, dispatcher, @global_state)
      end
    end
  end
end
