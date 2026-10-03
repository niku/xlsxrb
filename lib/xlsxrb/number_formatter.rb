# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"

module Xlsxrb
  # Formats cell values according to OpenXML formatCode strings.
  # Evaluates and converts numeric, currency, percentage, date/time, scientific,
  # and accounting format patterns into string representations.
  #
  # @api public
  module NumberFormatter
    # Built-in date and time format code strings.
    DATE_PATTERNS = [
      "mm-dd-yy", "d-mmm-yy", "d-mmm", "mmm-yy", "h:mm AM/PM", "h:mm:ss AM/PM", "h:mm", "h:mm:ss", "m/d/yy h:mm", "mm:ss", "[h]:mm:ss", "mmss.0"
    ].freeze

    class << self
      # Formats a cell value according to an Excel format code string.
      #
      # @param value [Object, nil] The cell's raw or typed value.
      # @param format_code [String, nil] Excel format code pattern.
      # @param date1904 [Boolean] Whether the 1904 date system is active.
      # @return [String] Formatted string representation.
      # @api public
      #: (untyped value, ?String? format_code, ?date1904: bool) -> String
      def format(value, format_code = nil, date1904: false)
        return "" if value.nil?
        return value.code.to_s if value.respond_to?(:code)
        return value ? "TRUE" : "FALSE" if [true, false].include?(value)

        code = format_code.to_s.strip
        return default_format(value) if code.empty? || code.casecmp("general").zero?

        return format_datetime(value, code) if value.is_a?(Date) || value.is_a?(DateTime) || value.is_a?(Time)

        if value.is_a?(Numeric) && date_format?(code)
          dt = if value.is_a?(Float) && (value % 1).positive?
                 Ooxml::Utils.serial_to_datetime(value, date1904: date1904)
               else
                 Ooxml::Utils.serial_to_date(value.to_i, date1904: date1904)
               end
          return format_datetime(dt, code)
        end

        return format_number(value, code) if value.is_a?(Numeric)

        if code.include?("@")
          sections = split_sections(code)
          pattern = sections[3] || sections[0]
          pattern.gsub("@", value.to_s)
        else
          value.to_s
        end
      end

      # Resolves the format code string for a given style index from a styles definition hash.
      #
      # @param style_index [Integer, nil] 0-based style index in cellXfs.
      # @param styles [Hash, nil] Parsed styles hash.
      # @return [String, nil] Format code string, or nil if not found.
      # @api public
      #: (Integer? style_index, Hash[untyped, untyped]? styles) -> String?
      def format_code_for(style_index, styles)
        return nil unless style_index && styles

        cell_xfs = styles[:cell_xfs]
        return nil unless cell_xfs.is_a?(Array) && style_index < cell_xfs.size

        xf = cell_xfs[style_index]
        return nil unless xf.is_a?(Hash)

        num_fmt_id = xf[:num_fmt_id]
        if (num_fmt_id.nil? || num_fmt_id.zero?) && xf[:xf_id] && styles[:cell_style_xfs].is_a?(Array)
          parent_xf = styles[:cell_style_xfs][xf[:xf_id]]
          num_fmt_id = parent_xf[:num_fmt_id] if parent_xf.is_a?(Hash)
        end

        return nil unless num_fmt_id

        num_fmts = styles[:num_fmts] || {}
        num_fmts[num_fmt_id] || Ooxml::Utils::BUILTIN_NUM_FMT_CODES[num_fmt_id]
      end

      # Checks whether a format code string represents a date or time pattern.
      #
      # @param format_code [String, nil]
      # @return [Boolean]
      #: (String? format_code) -> bool
      def date_format?(format_code)
        return false if format_code.nil? || format_code.empty?
        return true if DATE_PATTERNS.include?(format_code)

        stripped = format_code.gsub(/"[^"]*"/, "").gsub(/\[[^\]]*\]/, "").gsub(/\\[.]/, "")
        stripped.match?(/[ymdhsYMDHS]/)
      end

      private

      def default_format(value)
        if value.is_a?(Float)
          value == value.to_i ? value.to_i.to_s : value.to_s
        else
          value.to_s
        end
      end

      def split_sections(code)
        sections = []
        current = String.new
        in_quotes = false
        code.each_char do |ch|
          if ch == "\""
            in_quotes = !in_quotes
            current << ch
          elsif ch == ";" && !in_quotes
            sections << current
            current = String.new
          else
            current << ch
          end
        end
        sections << current
        sections
      end

      def format_number(value, code)
        sections = split_sections(code)
        if sections.size == 1
          pattern = sections[0]
          is_negative = value.negative?
          num = value.abs
        elsif sections.size >= 2
          if value.negative?
            pattern = sections[1]
            is_negative = false
            num = value.abs
          elsif value.zero? && sections.size >= 3
            pattern = sections[2]
            is_negative = false
            num = 0
          else
            pattern = sections[0]
            is_negative = false
            num = value
          end
        end

        return (is_negative ? "-#{num}" : num.to_s) if pattern == "@"

        format_single_number(num, pattern, is_negative)
      end

      def format_single_number(num, pattern, is_neg)
        clean = pattern.strip
        prefix = ""
        suffix = ""

        if clean =~ /\A\\-/ || clean =~ /\A-/
          prefix = "-"
          clean = clean.sub(/\A\\-/, "").sub(/\A-/, "")
        end

        clean = clean.gsub(/_[-_()€$]/, "").gsub(/\* */, "").gsub("\\ ", "").gsub(/"[^"]*"/, "").strip if clean.include?("_-") || clean.include?("*-") || clean.include?("* ")

        case clean
        when /\A(\[Red\])?\((.*)\)\z/
          prefix = "#{::Regexp.last_match(1)}("
          suffix = ")"
          clean = ::Regexp.last_match(2)
        when /\A"([^"]+)"(.*)\z/, /\A\[\$([^\]-]+)[^\]]*\]\s*(.*)\z/, /\A([$€£¥])(.*)\z/
          prefix = ::Regexp.last_match(1)
          clean = ::Regexp.last_match(2)
        end

        if clean =~ /0\.0+E\+00/i
          dec = clean[/0\.(0+)E/i, 1].size
          formatted = Kernel.format("%.#{dec}E", num)
          return "#{prefix}#{formatted}#{suffix}"
        elsif clean =~ /##0\.0+E\+0/i
          dec = clean[/##0\.(0+)E/i, 1].size
          formatted = Kernel.format("%.#{dec}E", num)
          return "#{prefix}#{formatted}#{suffix}"
        end

        if clean.end_with?("%")
          clean = clean.sub(/%+\z/, "")
          suffix = "%#{suffix}"
          num = num.to_f * 100
        end

        if clean =~ /\A(0+)\z/
          zeros = ::Regexp.last_match(1).size
          formatted = Kernel.format("%0#{zeros}d", num.round)
          return "#{"-" if is_neg}#{prefix}#{formatted}#{suffix}"
        end

        decimals = 0
        decimals = ::Regexp.last_match(1).size if clean =~ /\.(0+)/

        has_commas = clean.include?(",")

        str = Kernel.format("%.#{decimals}f", num)
        if has_commas
          parts = str.split(".")
          parts[0] = parts[0].reverse.gsub(/(\d{3})(?=\d)/, "\\1,").reverse
          str = parts.join(".")
        end

        "#{"-" if is_neg}#{prefix}#{str}#{suffix}"
      end

      def format_datetime(datetime, format_code)
        strftime_pattern = excel_to_strftime(format_code)
        datetime.strftime(strftime_pattern)
      rescue StandardError
        datetime.strftime("%Y-%m-%d %H:%M:%S")
      end

      def excel_to_strftime(format_code)
        code = format_code.dup

        code.gsub!(%r{am/pm}i, "%p")
        code.gsub!(%r{a/p}i, "%p")

        code.gsub!(/\[h+\]/i, "[%-k]")

        code.gsub!(/(h+:)(mm)/i, "\\1%M")
        code.gsub!(/(h+:)(m)(?!%)/i, "\\1%-M")
        code.gsub!(/(mm)(:s+)/i, "%M\\2")
        code.gsub!(/(?<!%)(m)(:s+)/i, "%-M\\2")

        code.gsub!(/ss/i, "%S")
        code.gsub!(/(?<!%)s/i, "%-S")

        code.gsub!("000", "%3N")
        code.gsub!("00", "%2N")
        code.gsub!("0", "%1N")

        code.gsub!(/mmmm/i, "%B")
        code.gsub!(/mmm/i, "%^b")
        code.gsub!(/mm/i, "%m")
        code.gsub!(/(?<!%)(?<![a-zA-Z])m(?![a-zA-Z])/i, "%-m")

        if code.include?("%p")
          code.gsub!(/hh/i, "%I")
          code.gsub!(/(?<!%)h/i, "%-l")
        else
          code.gsub!(/hh/i, "%H")
          code.gsub!(/(?<!%)h/i, "%-k")
        end

        code.gsub!(/dddd/i, "%A")
        code.gsub!(/ddd/i, "%^a")
        code.gsub!(/dd/i, "%d")
        code.gsub!(/(?<!%)d/i, "%-d")

        code.gsub!(/yyyy/i, "%Y")
        code.gsub!(/yy/i, "%y")

        code.gsub!(/\\(.)/, "\\1")

        code
      end
    end
  end
end
