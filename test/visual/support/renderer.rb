# frozen_string_literal: true

require "fileutils"
require "open3"
require "tmpdir"
require "securerandom"

module Xlsxrb
  module Visual
    class Renderer
      def self.soffice_command
        return ENV["SOFFICE_PATH"] if ENV["SOFFICE_PATH"] && !ENV["SOFFICE_PATH"].empty?

        # Try default PATH first
        return "soffice" if system("command -v soffice >/dev/null 2>&1")

        # Try common macOS paths
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        return mac_path if File.exist?(mac_path)

        # Try common Windows paths
        win_path = "C:/Program Files/LibreOffice/program/soffice.exe"
        return win_path if File.exist?(win_path)

        # Fallback to default
        "soffice"
      end

      def self.render(xlsx_path, output_dir)
        FileUtils.mkdir_p(output_dir)
        basename = File.basename(xlsx_path, ".xlsx")
        pdf_path = File.join(output_dir, "#{basename}.pdf")

        soffice_cmd = soffice_command
        raise Errno::ENOENT, "soffice (LibreOffice) binary not found. Please install LibreOffice or set the SOFFICE_PATH environment variable (e.g. SOFFICE_PATH=/Applications/LibreOffice.app/Contents/MacOS/soffice)." if soffice_cmd == "soffice" && !system("command -v soffice >/dev/null 2>&1")

        # 1. Convert XLSX to PDF using LibreOffice in headless mode with a unique clean user profile
        profile_dir = File.join(Dir.tmpdir, "libreoffice_user_profile_#{Process.pid}_#{SecureRandom.hex(6)}")
        begin
          stdout, stderr, status = Open3.capture3(
            soffice_cmd,
            "-env:UserInstallation=file://#{profile_dir}",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", output_dir,
            xlsx_path
          )

          raise "LibreOffice conversion failed: #{stderr}\n#{stdout}" unless status.success?
        ensure
          FileUtils.rm_rf(profile_dir)
        end

        # 2. Render PDF pages to PNG using pdftoppm
        # Generates page-1.png, page-2.png, etc.
        prefix = File.join(output_dir, "page")
        stdout, stderr, status = Open3.capture3(
          "pdftoppm",
          "-png",
          "-r", "150",
          pdf_path,
          prefix
        )

        raise "pdftoppm rendering failed: #{stderr}\n#{stdout}" unless status.success?

        # Clean up temporary PDF
        FileUtils.rm_f(pdf_path)

        # Return list of generated PNG paths sorted by page number
        Dir.glob(File.join(output_dir, "page-*.png")).sort_by do |path|
          path.match(/page-(\d+)\.png/)[1].to_i
        end
      end
    end
  end
end
