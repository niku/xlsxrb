# frozen_string_literal: true

require "test_helper"
require "tmpdir"

class EncryptionRoundtripTest < Test::Unit::TestCase
  def test_streaming_write_and_read_encrypted_file
    Dir.mktmpdir do |dir|
      file_path = File.join(dir, "encrypted_test.xlsx")
      password = "SecurePassword@2026"

      # 1. Streaming Write with Password
      Xlsxrb.write(file_path, password: password) do |wb|
        wb.sheet("Confidential") do |s|
          s.row(["Header 1", "Header 2", "Amount"])
          s.row(["Alice", "Department A", 15_000])
          s.row(["Bob", "Department B", 25_000])
        end
      end

      assert_true File.exist?(file_path)
      raw_file_bytes = File.binread(file_path)
      assert_true Xlsxrb::Ooxml::Crypto.encrypted?(raw_file_bytes)

      # 2. Reading without password must raise EncryptedFileError
      assert_raise(Xlsxrb::EncryptedFileError) do
        Xlsxrb.read(file_path)
      end

      # 3. Reading with wrong password must raise InvalidPasswordError
      assert_raise(Xlsxrb::InvalidPasswordError) do
        Xlsxrb.read(file_path, password: "WrongPassword")
      end

      # 4. Streaming Read with correct password
      collected_rows = []
      Xlsxrb.read(file_path, password: password) do |sheet|
        assert_equal "Confidential", sheet.name
        sheet.each_row do |row|
          collected_rows << row.cells.map(&:value)
        end
      end

      assert_equal 3, collected_rows.size
      assert_equal ["Header 1", "Header 2", "Amount"], collected_rows[0]
      assert_equal ["Alice", "Department A", 15_000], collected_rows[1]
      assert_equal ["Bob", "Department B", 25_000], collected_rows[2]

      # 5. In-memory load with correct password
      wb = Xlsxrb.read(file_path, password: password).load
      sheet = wb.sheets.first
      assert_equal "Confidential", sheet.name
      assert_equal "Alice", sheet["A2"].value
      assert_equal 25_000, sheet["C3"].value
    end
  end

  def test_in_memory_write_and_read_encrypted_binary_string
    password = "MyInMemPassword!"

    wb = Xlsxrb.build do |b|
      b.sheet("Sales") do |s|
        s.row(%w[Product Price])
        s.row(["Laptop", 1200.5])
      end
    end

    # Export to encrypted binary string
    encrypted_bytes = Xlsxrb.write(wb, password: password)
    assert_true Xlsxrb::Ooxml::Crypto.encrypted?(encrypted_bytes)

    # Read from StringIO with password
    read_wb = Xlsxrb.read(StringIO.new(encrypted_bytes), password: password).load
    assert_equal 1, read_wb.sheets.size
    sales_sheet = read_wb.sheets[0]
    assert_equal "Sales", sales_sheet.name
    assert_equal "Laptop", sales_sheet["A2"].value
    assert_equal 1200.5, sales_sheet["B2"].value
  end

  def test_modify_encrypted_file
    Dir.mktmpdir do |dir|
      file_path = File.join(dir, "modify_test.xlsx")
      password = "SecretModifyKey"

      Xlsxrb.write(file_path, password: password) do |wb|
        wb.sheet("Data") do |s|
          s.row(["Original Value"])
        end
      end

      # Modify file with password
      Xlsxrb.modify(file_path, password: password) do |workbook|
        workbook.update_sheet("Data") do |sheet|
          sheet.update_cell("A1", value: "Updated Secure Value")
        end
      end

      # Verify modified content
      assert_true Xlsxrb::Ooxml::Crypto.encrypted?(File.binread(file_path))
      assert_raise(Xlsxrb::EncryptedFileError) { Xlsxrb.read(file_path) }
      assert_raise(Xlsxrb::InvalidPasswordError) { Xlsxrb.read(file_path, password: "wrong").load }

      reloaded_wb = Xlsxrb.read(file_path, password: password).load
      assert_equal "Updated Secure Value", reloaded_wb.sheets[0]["A1"].value
    end
  end
end
