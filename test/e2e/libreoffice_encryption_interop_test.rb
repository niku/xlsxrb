# frozen_string_literal: true

require "test_helper"
require "tmpdir"
require "open3"

module Xlsxrb
  module Ooxml
    class LibreofficeEncryptionInteropTest < Test::Unit::TestCase
      def test_libreoffice_can_decrypt_and_render_encrypted_xlsx
        return unless system("which soffice > /dev/null 2>&1") && system("python3 -c 'import uno' > /dev/null 2>&1")

        validator_py = File.expand_path("../fixtures/lo_uno_validator.py", __dir__)
        return unless File.exist?(validator_py)

        Dir.mktmpdir do |dir|
          xlsx_path = File.join(dir, "lo_encrypted.xlsx")
          password = "LibreOfficePass2026"

          # Generate encrypted file with xlsxrb
          Xlsxrb.write(xlsx_path, password: password) do |wb|
            wb.sheet("EncryptedData") do |s|
              s.row(%w[Title Score])
              s.row(["Alpha", 100])
              s.row(["Beta", 200])
            end
          end

          assert_true File.exist?(xlsx_path)
          assert_true Crypto.encrypted?(File.binread(xlsx_path))

          # 1. Attempt to open with LibreOffice WITHOUT password (or wrong password) -> should fail (exit code != 0)
          _stdout_fail, _stderr_fail, status_fail = Open3.capture3(
            "python3", validator_py,
            "--file", xlsx_path,
            "--password", "WrongPassword"
          )
          assert_false status_fail.success?

          # 2. Open with LibreOffice WITH correct password -> should succeed (exit code == 0)
          stdout_ok, stderr_ok, status_ok = Open3.capture3(
            "python3", validator_py,
            "--file", xlsx_path,
            "--password", password
          )

          assert_true status_ok.success?, "LibreOffice UNO failed to decrypt and open file: #{stdout_ok} #{stderr_ok}"
          assert_includes stdout_ok, "A1=Title"
          assert_includes stdout_ok, "B1=Score"
        end
      end

      def test_xlsxrb_can_read_encrypted_file_created_by_libreoffice
        return unless system("which soffice > /dev/null 2>&1") && system("python3 -c 'import uno' > /dev/null 2>&1")

        Dir.mktmpdir do |dir|
          lo_xlsx_path = File.join(dir, "lo_created.xlsx")
          password = "CreatedByLibreOffice123"

          # Use LibreOffice Python UNO to create a password protected XLSX
          create_py = <<~PY
            import uno
            from com.sun.star.beans import PropertyValue
            import subprocess, time, os

            port = 2020 + (os.getpid() % 1000)
            lo_proc = subprocess.Popen([
                "soffice", "--headless", f"--accept=socket,host=127.0.0.1,port={port};urp;",
                "-env:UserInstallation=file:///tmp/lo_create_p_{}".format(os.getpid())
            ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            try:
                local_ctx = uno.getComponentContext()
                smgr = local_ctx.ServiceManager
                resolver = smgr.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local_ctx)
                ctx = None
                for _ in range(30):
                    try:
                        ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
                        break
                    except Exception:
                        time.sleep(0.2)
                desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
                p_hidden = PropertyValue("Hidden", 0, True, 0)
                doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (p_hidden,))
                sheet = doc.getSheets().getByIndex(0)
                sheet.getCellByPosition(0, 0).setString("FromLibreOffice")
                sheet.getCellByPosition(1, 0).setValue(9876.5)

                p_filter = PropertyValue("FilterName", 0, "Calc Office Open XML", 0)
                p_pass = PropertyValue("Password", 0, "#{password}", 0)
                doc.storeToURL(uno.systemPathToFileUrl("#{lo_xlsx_path}"), (p_filter, p_pass))
                doc.close(True)
            finally:
                lo_proc.terminate()
                lo_proc.wait()
          PY

          system("python3", "-c", create_py)
          assert_true File.exist?(lo_xlsx_path)

          # Verify xlsxrb fails without password
          assert_raise(Xlsxrb::EncryptedFileError) do
            Xlsxrb.read(lo_xlsx_path)
          end

          # Verify xlsxrb fails with wrong password
          assert_raise(Xlsxrb::InvalidPasswordError) do
            Xlsxrb.read(lo_xlsx_path, password: "BadPassword")
          end

          # Verify xlsxrb successfully reads with correct password
          wb = Xlsxrb.read(lo_xlsx_path, password: password).load
          assert_equal 1, wb.sheets.size
          sheet = wb.sheets[0]
          assert_equal "FromLibreOffice", sheet["A1"].value
          assert_equal 9876.5, sheet["B1"].value
        end
      end
    end
  end
end
