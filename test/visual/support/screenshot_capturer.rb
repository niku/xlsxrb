# frozen_string_literal: true

require "fileutils"
require "tmpdir"

module Xlsxrb
  module Visual
    module ScreenshotCapturer
      # Configuration for interactive examples
      INTERACTIVE_CONFIGS = {
        "interactive_autofilter" => {
          xlsx: "docs/visual/files/interactive_autofilter.xlsx",
          png: "test/visual/support/illustrations/interactive_autofilter_page-2.png",
          actions: [
            { type: :key, value: "alt+Down" }
          ]
        },
        "interactive_validation_list" => {
          xlsx: "docs/visual/files/interactive_validation_list.xlsx",
          png: "test/visual/support/illustrations/interactive_validation_list_page-2.png",
          actions: [
            { type: :key, value: "Right" },
            { type: :key, value: "Down" },
            { type: :key, value: "alt+Down" }
          ]
        },
        "interactive_validation_range" => {
          xlsx: "docs/visual/files/interactive_validation_range.xlsx",
          png: "test/visual/support/illustrations/interactive_validation_range_page-2.png",
          actions: [
            { type: :key, value: "Right" },
            { type: :key, value: "Down" },
            { type: :type, value: "5" },
            { type: :key, value: "Return" }
          ]
        },
        "interactive_comments" => {
          xlsx: "docs/visual/files/interactive_comments.xlsx",
          png: "test/visual/support/illustrations/interactive_comments_page-2.png",
          actions: [
            { type: :key, value: "Shift+F10" },
            { type: :key, value: "w" }
          ]
        },
        "interactive_validation_date" => {
          xlsx: "docs/visual/files/interactive_validation_date.xlsx",
          png: "test/visual/support/illustrations/interactive_validation_date_page-2.png",
          actions: [
            { type: :key, value: "Right" },
            { type: :key, value: "Down" },
            { type: :type, value: "2025-01-01" },
            { type: :key, value: "Return" }
          ]
        },
        "interactive_validation_text_length" => {
          xlsx: "docs/visual/files/interactive_validation_text_length.xlsx",
          png: "test/visual/support/illustrations/interactive_validation_text_length_page-2.png",
          actions: [
            { type: :key, value: "Right" },
            { type: :key, value: "Down" },
            { type: :type, value: "ThisTextIsTooLong" },
            { type: :key, value: "Return" }
          ]
        },
        "interactive_validation_custom" => {
          xlsx: "docs/visual/files/interactive_validation_custom.xlsx",
          png: "test/visual/support/illustrations/interactive_validation_custom_page-2.png",
          actions: [
            { type: :key, value: "Right" },
            { type: :key, value: "Down" },
            { type: :type, value: "5" },
            { type: :key, value: "Return" }
          ]
        },
        "interactive_validation_time" => {
          xlsx: "docs/visual/files/interactive_validation_time.xlsx",
          png: "test/visual/support/illustrations/interactive_validation_time_page-2.png",
          actions: [
            { type: :key, value: "Right" },
            { type: :key, value: "Down" },
            { type: :type, value: "07:00" },
            { type: :key, value: "Return" }
          ]
        },
        "sheet_tab_colors" => {
          xlsx: "docs/visual/files/sheet_tab_colors.xlsx",
          png: "test/visual/support/illustrations/sheet_tab_colors_page-2.png",
          actions: []
        },
        "sparkline_line" => {
          xlsx: "docs/visual/files/sparkline_line.xlsx",
          png: "test/visual/support/illustrations/sparkline_line_page-2.png",
          actions: []
        }
      }.freeze

      def self.tools_available?
        system("which xvfb-run >/dev/null 2>&1") &&
          system("which xdotool >/dev/null 2>&1") &&
          system("which scrot >/dev/null 2>&1") &&
          system("which soffice >/dev/null 2>&1")
      end

      def self.capture_all
        unless tools_available?
          puts "Xvfb, xdotool, scrot, or LibreOffice not available. Skipping interactive screenshots generation."
          return
        end

        puts "Tools detected. Generating interactive screenshots..."
        INTERACTIVE_CONFIGS.each do |name, config|
          xlsx_path = File.expand_path("../../../#{config[:xlsx]}", __dir__)
          png_path = File.expand_path("../../../#{config[:png]}", __dir__)

          # Ensure output directory exists
          FileUtils.mkdir_p(File.dirname(png_path))

          puts "Capturing GUI state for #{name}..."
          capture_single(xlsx_path, png_path, config[:actions])
        end
        true
      end

      def self.capture_single(xlsx_path, png_path, actions)
        # Start LibreOffice Calc under Xvfb
        # We wrapper the entire ruby subprocess or start soffice inside xvfb-run
        # To avoid nested xvfb-run issues, we just run soffice directly
        # and assume the caller runs us inside xvfb-run, or we run soffice via xvfb-run.
        # Running soffice via xvfb-run works if we set DISPLAY inside xvfb-run.
        # But wait! If we run `xvfb-run --server-args="-screen 0 1024x768x24" soffice`,
        # then only soffice is inside Xvfb. Since we need xdotool and scrot to share the same DISPLAY,
        # we can wrap the single execution block under a nested xvfb-run!

        # Let's write a temporary shell script to run the actions inside xvfb-run
        script_file = File.join(Dir.tmpdir, "screenshot_actions.sh")
        File.open(script_file, "w") do |f|
          f.puts "#!/bin/bash"
          f.puts "soffice --norestore --nologo --calc \"#{xlsx_path}\" &"
          f.puts "SOFFICE_PID=$!"
          f.puts "sleep 6"
          f.puts "WIN_ID=$(xdotool search --name \"LibreOffice Calc\" | head -n 1)"
          f.puts "if [ ! -z \"$WIN_ID\" ]; then"
          f.puts "  xdotool windowfocus --sync \"$WIN_ID\""
          f.puts "  sleep 1"
          # Write all actions
          actions.each do |act|
            if act[:type] == :key
              f.puts "  xdotool key \"#{act[:value]}\""
            elsif act[:type] == :type
              f.puts "  xdotool type \"#{act[:value]}\""
            end
            f.puts "  sleep 1"
          end
          f.puts "fi"
          f.puts "sleep 1"
          f.puts "scrot -o \"#{png_path}\""
          f.puts "kill $SOFFICE_PID"
          f.puts "killall -9 soffice.bin 2>/dev/null || true"
          f.puts "wait $SOFFICE_PID 2>/dev/null || true"
        end
        FileUtils.chmod(0o755, script_file)

        # Run the script under xvfb-run
        system("xvfb-run --auto-servernum --server-args=\"-screen 0 1024x768x24\" #{script_file}")
      ensure
        FileUtils.rm_f(script_file) if script_file
      end
    end
  end
end
