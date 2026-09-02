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
    end
  end
end
