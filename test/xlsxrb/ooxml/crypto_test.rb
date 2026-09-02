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
    end
  end
end
