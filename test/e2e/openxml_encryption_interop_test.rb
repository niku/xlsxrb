# frozen_string_literal: true

require "test_helper"

class OpenxmlEncryptionInteropTest < Test::Unit::TestCase
  def test_openxml_sdk_can_validate_xlsxrb_encrypted_file
    return unless system("which dotnet > /dev/null 2>&1")

    Dir.mktmpdir do |dir|
      xlsx_path = File.join(dir, "encrypted_for_sdk.xlsx")
      password = "OpenXmlSecretPass123"

      # Generate encrypted XLSX with xlsxrb
      Xlsxrb.write(xlsx_path, password: password) do |wb|
        wb.sheet("ValidatedSheet") do |s|
          s.row(%w[Product Qty Price])
          s.row(["Widget A", 10, 29.99])
          s.row(["Gadget B", 5, 99.50])
        end
      end

      assert_true File.exist?(xlsx_path)
      assert_true Xlsxrb::Ooxml::Crypto.encrypted?(File.binread(xlsx_path))

      # Run OpenXml SDK validation scenario
      scenario_path = File.expand_path("../fixtures/sdk_scenarios/encryption_standard_validation.cs", __dir__)
      result = OpenXmlSdkScenarioRunner.run_single_scenario(scenario_path, xlsx_path)

      assert_true result[:success], "OpenXml SDK validation failed: #{result[:stderr]}\n#{result[:stdout]}"
      assert_includes result[:stdout], "SCENARIO_PASS"
    end
  end

  def test_xlsxrb_can_encrypt_and_read_openxml_sdk_generated_xlsx
    return unless system("which dotnet > /dev/null 2>&1")

    scenario_path = File.expand_path("../fixtures/sdk_scenarios/basic_sheet_generated_by_sdk.cs", __dir__)
    return unless File.exist?(scenario_path)

    Dir.mktmpdir do |dir|
      sdk_xlsx_path = File.join(dir, "from_sdk.xlsx")
      enc_xlsx_path = File.join(dir, "from_sdk_encrypted.xlsx")
      password = "SdkRoundtripPass456"

      result = OpenXmlSdkScenarioRunner.run_single_scenario(scenario_path, sdk_xlsx_path)
      assert_true result[:success], "Failed to generate SDK fixture: #{result[:stderr]}"

      # Encrypt SDK-generated XLSX with xlsxrb
      plain_bytes = File.binread(sdk_xlsx_path)
      encrypted_bytes = Xlsxrb::Ooxml::Crypto.encrypt(plain_bytes, password)
      File.binwrite(enc_xlsx_path, encrypted_bytes)

      assert_true Xlsxrb::Ooxml::Crypto.encrypted?(File.binread(enc_xlsx_path))

      # Decrypt and read with Xlsxrb
      wb = Xlsxrb.read(enc_xlsx_path, password: password).load
      assert_not_nil wb
      assert_operator wb.sheets.size, :>, 0
    end
  end
end
