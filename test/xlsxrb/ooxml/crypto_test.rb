# frozen_string_literal: true

require "test_helper"
require_relative "../../../lib/xlsxrb/ooxml/crypto"

module Xlsxrb
  module Ooxml
    class CryptoTest < Test::Unit::TestCase
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
      end
    end
  end
end
