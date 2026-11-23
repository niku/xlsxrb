# frozen_string_literal: true

require "test_helper"
require_relative "../../../lib/xlsxrb/ooxml/crypto"

module Xlsxrb
  module Ooxml
    class CryptoTest < Test::Unit::TestCase
      cover Xlsxrb::Ooxml::Crypto

      def test_encrypt_and_decrypt_roundtrip
        plain_payload = "PK\x03\x04Hello this is a fake zip payload for testing encryption #{"A" * 10_000}".b
        password = "secret_password_123"

        encrypted_cfb = Crypto.encrypt(plain_payload, password)
        assert_true Crypto.encrypted?(encrypted_cfb)

        # Decrypt with correct password
        decrypted = Crypto.decrypt(encrypted_cfb, password)
        assert_equal plain_payload, decrypted
      end

      def test_decrypt_with_invalid_password
        plain_payload = "PK\x03\x04Hello world".b
        password = "correct_password"

        encrypted_cfb = Crypto.encrypt(plain_payload, password)

        assert_raise(Xlsxrb::InvalidPasswordError) do
          Crypto.decrypt(encrypted_cfb, "wrong_password")
        end
      end

      def test_decrypt_without_password
        plain_payload = "PK\x03\x04Hello world".b
        password = "correct_password"

        encrypted_cfb = Crypto.encrypt(plain_payload, password)

        assert_raise(Xlsxrb::EncryptedFileError) do
          Crypto.decrypt(encrypted_cfb, nil)
        end
      end

      def test_encrypted_predicate_on_non_encrypted
        plain_zip = "PK\x03\x04Hello world".b
        assert_false Crypto.encrypted?(plain_zip)
        assert_false Crypto.encrypted?("short".b)
        assert_false Crypto.encrypted?("")
      end

      def test_encrypt_and_decrypt_agile_mode_roundtrip
        plain_payload = "PK\x03\x04Agile encryption test payload #{"B" * 5000}".b
        password = "agile_secret_456"

        encrypted_cfb = Crypto.encrypt(plain_payload, password, mode: :agile)
        assert_true Crypto.encrypted?(encrypted_cfb)

        decrypted = Crypto.decrypt(encrypted_cfb, password)
        assert_equal plain_payload, decrypted
      end

      # Checks the HMAC as [MS-OFFCRYPTO] 2.3.4.14 describes it, apart from the reader, which
      # also accepts the HMAC that xlsxrb 0.2.1 and earlier wrote.
      def test_agile_writes_the_hmac_of_the_encrypted_package
        password = "agile_secret_456"
        streams = Crypto::Agile.encrypt("PK\x03\x04#{"B" * 5000}".b, password)
        package = streams["EncryptedPackage"]
        doc = REXML::Document.new(streams["EncryptionInfo"][8..])
        key_data = doc.root.elements["keyData"].attributes
        encryptor = doc.root.elements["keyEncryptors/keyEncryptor/*"].attributes
        integrity = doc.root.elements["dataIntegrity"].attributes
        decrypt = lambda do |key, init_vector, data|
          cipher = OpenSSL::Cipher.new("aes-256-cbc").decrypt
          cipher.key = key
          cipher.iv = init_vector[0, 16]
          cipher.padding = 0
          cipher.update(data) + cipher.final
        end

        hash = OpenSSL::Digest::SHA512.digest(encryptor["saltValue"].unpack1("m0") + password.encode("UTF-16LE").b)
        encryptor["spinCount"].to_i.times { |i| hash = OpenSSL::Digest::SHA512.digest([i].pack("V") + hash) }
        key_key = OpenSSL::Digest::SHA512.digest(hash + Crypto::Agile::BLOCK_KEY_KEY)[0, 32]
        package_key = decrypt.call(key_key, encryptor["saltValue"].unpack1("m0"), encryptor["encryptedKeyValue"].unpack1("m0"))[0, 32]
        salt = key_data["saltValue"].unpack1("m0")
        hmac_key = decrypt.call(package_key, OpenSSL::Digest::SHA512.digest(salt + Crypto::Agile::BLOCK_KEY_INTEGRITY_KEY), integrity["encryptedHmacKey"].unpack1("m0"))
        hmac_value = decrypt.call(package_key, OpenSSL::Digest::SHA512.digest(salt + Crypto::Agile::BLOCK_KEY_INTEGRITY_VALUE), integrity["encryptedHmacValue"].unpack1("m0"))

        assert_equal 64, hmac_key.bytesize
        assert_equal OpenSSL::HMAC.digest("SHA512", hmac_key, package), hmac_value
      end

      def test_reads_files_encrypted_by_apache_poi
        { "poi_agile.xlsx" => "Poi-Passw0rd-agile", "poi_standard.xlsx" => "Poi-Passw0rd-standard" }.each do |name, password|
          path = encryption_fixture(name)
          workbook = Xlsxrb.load(path, password: password)

          assert_equal "plain", workbook.sheets[0]["A1"].value, name
          assert_equal "y", workbook.sheets[0]["B2"].value, name
          assert_raise(Xlsxrb::InvalidPasswordError, name) { Xlsxrb.load(path, password: "wrong") }
        end
      end

      def test_agile_rejects_a_changed_package
        info, package = encryption_streams("poi_agile.xlsx")

        [0, 8, package.bytesize / 2, package.bytesize - 1].each do |position|
          changed = package.dup
          changed.setbyte(position, changed.getbyte(position) ^ 1)

          assert_raise_with_message(Xlsxrb::DecryptionError, /Data integrity check failed/, "byte #{position}") do
            Crypto::Agile.decrypt(info, changed, "Poi-Passw0rd-agile")
          end
        end
        assert_raise_with_message(Xlsxrb::DecryptionError, /Data integrity check failed/) do
          Crypto::Agile.decrypt(info, "#{package}#{"\0" * 16}", "Poi-Passw0rd-agile")
        end
      end

      def test_agile_rejects_encryption_info_without_data_integrity
        info, package = encryption_streams("poi_agile.xlsx")
        stripped = info.sub(%r{<dataIntegrity[^>]*/>}, "")

        assert_not_equal info, stripped
        assert_raise_with_message(Xlsxrb::DecryptionError, /Missing dataIntegrity/) do
          Crypto::Agile.decrypt(stripped, package, "Poi-Passw0rd-agile")
        end
      end

      # xlsxrb 0.2.1 wrote this file. Its HMAC is not the one the specification describes.
      def test_agile_reads_and_checks_files_written_before_the_hmac_fix
        workbook = Xlsxrb.load(encryption_fixture("xlsxrb_agile_legacy_hmac.xlsx"), password: "legacy-password")

        assert_equal "written before the HMAC fix", workbook.sheets[0]["A1"].value

        info, package = encryption_streams("xlsxrb_agile_legacy_hmac.xlsx")
        changed = package.dup
        changed.setbyte(package.bytesize / 2, changed.getbyte(package.bytesize / 2) ^ 1)

        assert_raise_with_message(Xlsxrb::DecryptionError, /Data integrity check failed/) do
          Crypto::Agile.decrypt(info, changed, "legacy-password")
        end
      end

      # Some writers leave bytes after the encrypted package that are not a whole AES block.
      def test_standard_ignores_bytes_after_the_package
        plain = "PK\x03\x04#{"S" * 1000}".b
        streams = Crypto::Standard.encrypt(plain, "pwd")
        info = streams["EncryptionInfo"]
        package = streams["EncryptedPackage"]

        assert_equal plain, Crypto::Standard.decrypt(info, "#{package}#{"\x01" * 8}".b, "pwd")
        assert_raise_with_message(Xlsxrb::DecryptionError, /shorter than its declared size/) do
          Crypto::Standard.decrypt(info, package[0, package.bytesize - 16], "pwd")
        end
      end

      def test_encrypt_with_nil_or_empty_password_returns_plain
        plain = "PK\x03\x04plain".b
        assert_equal plain, Crypto.encrypt(plain, nil)
        assert_equal plain, Crypto.encrypt(plain, "")
      end

      def test_decrypt_with_empty_password_raises
        plain = "PK\x03\x04plain".b
        encrypted = Crypto.encrypt(plain, "pwd")
        assert_raise(Xlsxrb::EncryptedFileError) do
          Crypto.decrypt(encrypted, "")
        end
      end

      def test_decrypt_with_missing_streams_raises
        # CFB with only dummy stream
        cfb = Cfb::Writer.write({ "Dummy" => "data".b })
        assert_raise(Xlsxrb::DecryptionError) do
          Crypto.decrypt(cfb, "pwd")
        end
      end

      def test_decrypt_with_unsupported_version_raises
        # EncryptionInfo with version 9.9
        unsupported_info = [9, 9].pack("vv") + ("\x00" * 32).b
        cfb = Cfb::Writer.write({
                                  "EncryptionInfo" => unsupported_info,
                                  "EncryptedPackage" => ("\x00" * 64).b
                                })
        assert_raise(Xlsxrb::DecryptionError) do
          Crypto.decrypt(cfb, "pwd")
        end
      end

      def test_standard_decrypt_package_bounds_and_corruption
        # Valid standard encrypted file
        plain = "PK\x03\x04test".b
        valid_cfb = Crypto.encrypt(plain, "password")
        reader = Cfb::Reader.new(valid_cfb)
        valid_info = reader.read_stream("EncryptionInfo")

        # Package stream too short (< 8 bytes)
        cfb_short_pkg = Cfb::Writer.write({
                                            "EncryptionInfo" => valid_info,
                                            "EncryptedPackage" => "short".b
                                          })
        assert_raise_with_message(Xlsxrb::DecryptionError, /Encrypted package stream is too short/) do
          Crypto.decrypt(cfb_short_pkg, "password")
        end

        # Package total_size exceeds 16GB limit
        huge_size_pkg = [0x500_000_000].pack("Q<") + ("\x00" * 32).b
        cfb_huge_pkg = Cfb::Writer.write({
                                           "EncryptionInfo" => valid_info,
                                           "EncryptedPackage" => huge_size_pkg
                                         })
        assert_raise_with_message(Xlsxrb::DecryptionError, /Encrypted package size is invalid or exceeds limits/) do
          Crypto.decrypt(cfb_huge_pkg, "password")
        end
      end

      def test_standard_parse_encryption_info_corrupt_offsets
        pkg = [4].pack("Q<") + ("\x00" * 32).b

        # Info stream too short (< 40 bytes, but identified as standard encryption minor=2)
        short_info_bytes = [3, 2].pack("vv") + ("\x00" * 26).b
        cfb_short_info = Cfb::Writer.write({
                                             "EncryptionInfo" => short_info_bytes,
                                             "EncryptedPackage" => pkg
                                           })
        assert_raise_with_message(Xlsxrb::DecryptionError, /EncryptionInfo stream too short/) do
          Crypto.decrypt(cfb_short_info, "password")
        end

        # Unsupported major/minor version (e.g. major 1, minor 2)
        bad_ver_info = [1, 2].pack("vv") + ("\x00" * 60).b
        cfb_bad_ver = Cfb::Writer.write({
                                          "EncryptionInfo" => bad_ver_info,
                                          "EncryptedPackage" => pkg
                                        })
        assert_raise_with_message(Xlsxrb::DecryptionError, /Unsupported standard encryption version/) do
          Crypto.decrypt(cfb_bad_ver, "password")
        end

        # Header offset out of bounds (large header_size)
        large_header_info = [3, 2, 0].pack("vvV") + [10_000].pack("V") + ("\x00" * 60).b
        cfb_large_header = Cfb::Writer.write({
                                               "EncryptionInfo" => large_header_info,
                                               "EncryptedPackage" => pkg
                                             })
        assert_raise_with_message(Xlsxrb::DecryptionError, /EncryptionInfo header offset out of bounds/) do
          Crypto.decrypt(cfb_large_header, "password")
        end

        # Salt size > 64 bytes
        # header_size = 32, verifier_offset = 44. At offset 44, put salt_size = 100
        salt_overflow_info = [3, 2, 0].pack("vvV") + [32].pack("V") + ("\x00" * 32).b + [100].pack("V") + ("\x00" * 128).b
        cfb_salt_overflow = Cfb::Writer.write({
                                                "EncryptionInfo" => salt_overflow_info,
                                                "EncryptedPackage" => pkg
                                              })
        assert_raise_with_message(Xlsxrb::DecryptionError, /Invalid salt size in Standard encryption/) do
          Crypto.decrypt(cfb_salt_overflow, "password")
        end
      end

      private

      def encryption_fixture(name)
        File.expand_path("../../fixtures/encryption/#{name}", __dir__)
      end

      def encryption_streams(name)
        reader = Cfb::Reader.new(File.binread(encryption_fixture(name)))
        [reader.read_stream("EncryptionInfo"), reader.read_stream("EncryptedPackage")]
      end
    end
  end
end
