# frozen_string_literal: true

# rbs_inline: enabled

require "openssl"
require "rexml/document"
require "securerandom"

module Xlsxrb
  module Ooxml
    module Crypto
      # Implements Microsoft Office Agile Encryption specified in [MS-OFFCRYPTO] Section 2.3.4.
      class Agile
        BLOCK_KEY_VERIFIER_INPUT   = "\xfe\xa7\xd2\x76\x3b\x4b\x9e\x79".b.freeze
        BLOCK_KEY_VERIFIER_VALUE   = "\xd7\xaa\x0f\x6d\x30\x61\x34\x4e".b.freeze
        BLOCK_KEY_KEY              = "\x14\x6e\x0b\xe7\xab\xac\xd0\xd6".b.freeze
        BLOCK_KEY_INTEGRITY_KEY    = "\x5f\xb2\xad\x01\x0c\xb9\xe1\xf6".b.freeze
        BLOCK_KEY_INTEGRITY_VALUE  = "\xa0\x67\x7f\x02\xb2\x2c\x84\x33".b.freeze

        SEGMENT_SIZE = 4096

        class << self
          # Decrypts an encrypted package given the EncryptionInfo stream data and EncryptedPackage stream data.
          def decrypt(encryption_info_bytes, encrypted_package_bytes, password)
            info = parse_encryption_info(encryption_info_bytes)
            password_str = password.to_s

            digest = digest_for(info[:hash_algorithm])
            cipher_name = cipher_name_for(info[:cipher_algorithm], info[:key_bits])

            # 1. Derive intermediate hash H_final from password and keyEncryptor salt
            h_final = derive_h_final(digest, password_str, info[:encryptor_salt], info[:spin_count])

            # 2. Verify password with verifier input and hash
            key_bytes = info[:key_bits] / 8
            k_ver_input = derive_key(digest, h_final, BLOCK_KEY_VERIFIER_INPUT, key_bytes)
            verifier_input = aes_decrypt(cipher_name, k_ver_input, info[:encryptor_salt], info[:encrypted_verifier_hash_input])

            expected_hash = digest.digest(verifier_input)

            k_ver_val = derive_key(digest, h_final, BLOCK_KEY_VERIFIER_VALUE, key_bytes)
            decrypted_hash = aes_decrypt(cipher_name, k_ver_val, info[:encryptor_salt], info[:encrypted_verifier_hash_value])

            hash_size = info[:hash_size]
            raise Xlsxrb::InvalidPasswordError, "Password verification failed" unless OpenSSL.secure_compare(decrypted_hash[0, hash_size], expected_hash[0, hash_size])

            # 3. Decrypt package master key
            k_key = derive_key(digest, h_final, BLOCK_KEY_KEY, key_bytes)
            package_key = aes_decrypt(cipher_name, k_key, info[:encryptor_salt], info[:encrypted_key_value])[0, key_bytes]

            # 4. Decrypt package stream
            raise Xlsxrb::DecryptionError, "Encrypted package stream is too short" if encrypted_package_bytes.bytesize < 8

            total_size = encrypted_package_bytes[0, 8].unpack1("Q<")
            raise Xlsxrb::DecryptionError, "Encrypted package size is invalid or exceeds limits" if total_size.negative? || total_size > 0x400_000_000

            encrypted_data = encrypted_package_bytes[8..] || "".b

            pkg_cipher_name = cipher_name_for(info[:key_data_cipher_algorithm], info[:key_data_key_bits])
            pkg_digest = digest_for(info[:key_data_hash_algorithm])
            pkg_salt = info[:key_data_salt]

            decrypted = +""
            num_blocks = (encrypted_data.bytesize + SEGMENT_SIZE - 1) / SEGMENT_SIZE
            num_blocks.times do |i|
              chunk = encrypted_data[i * SEGMENT_SIZE, SEGMENT_SIZE]
              break if chunk.nil? || chunk.empty?

              iv = pkg_digest.digest(pkg_salt + [i].pack("V"))[0, 16]
              decrypted << aes_decrypt_block(pkg_cipher_name, package_key, iv, chunk)
            end

            decrypted[0, total_size] || "".b
          end

          # Encrypts a plain zip payload with the given password into EncryptionInfo and EncryptedPackage streams.
          def encrypt(plain_bytes, password)
            password_str = password.to_s
            spin_count = 100_000
            key_bits = 256
            key_bytes = key_bits / 8
            hash_algorithm = "SHA512"
            cipher_algorithm = "AES"
            cipher_chaining = "ChainingModeCBC"
            hash_size = 64
            block_size = 16

            digest = digest_for(hash_algorithm)
            cipher_name = "aes-256-cbc"

            encryptor_salt = SecureRandom.random_bytes(16)
            key_data_salt = SecureRandom.random_bytes(16)
            package_key = SecureRandom.random_bytes(key_bytes)
            verifier_input = SecureRandom.random_bytes(16)

            # Derive intermediate hash H_final
            h_final = derive_h_final(digest, password_str, encryptor_salt, spin_count)

            # Encrypt verifier input
            k_ver_input = derive_key(digest, h_final, BLOCK_KEY_VERIFIER_INPUT, key_bytes)
            encrypted_verifier_input = aes_encrypt(cipher_name, k_ver_input, encryptor_salt, verifier_input)

            # Encrypt verifier hash value
            expected_hash = digest.digest(verifier_input)
            k_ver_val = derive_key(digest, h_final, BLOCK_KEY_VERIFIER_VALUE, key_bytes)
            encrypted_verifier_val = aes_encrypt(cipher_name, k_ver_val, encryptor_salt, expected_hash)

            # Encrypt package master key
            k_key = derive_key(digest, h_final, BLOCK_KEY_KEY, key_bytes)
            encrypted_key_value = aes_encrypt(cipher_name, k_key, encryptor_salt, package_key)

            # Encrypt plain package data in 4096-byte segments
            encrypted_package_stream = +""
            encrypted_package_stream << [plain_bytes.bytesize].pack("Q<")

            num_blocks = (plain_bytes.bytesize + SEGMENT_SIZE - 1) / SEGMENT_SIZE
            num_blocks.times do |i|
              chunk = plain_bytes[i * SEGMENT_SIZE, SEGMENT_SIZE] || "".b
              # Pad last chunk to multiple of block_size if necessary
              if (chunk.bytesize % block_size) != 0
                pad_len = block_size - (chunk.bytesize % block_size)
                chunk += ("\x00".b * pad_len)
              end

              iv = digest.digest(key_data_salt + [i].pack("V"))[0, 16]
              encrypted_package_stream << aes_encrypt_block(cipher_name, package_key, iv, chunk)
            end

            # Data integrity HMAC
            hmac_key = SecureRandom.random_bytes(64)
            hmac_value = OpenSSL::HMAC.digest(hash_algorithm, hmac_key, plain_bytes)

            k_integ_key = derive_key(digest, h_final, BLOCK_KEY_INTEGRITY_KEY, key_bytes)
            encrypted_hmac_key = aes_encrypt(cipher_name, k_integ_key, encryptor_salt, hmac_key)

            k_integ_val = derive_key(digest, h_final, BLOCK_KEY_INTEGRITY_VALUE, key_bytes)
            encrypted_hmac_val = aes_encrypt(cipher_name, k_integ_val, encryptor_salt, hmac_value)

            # Build EncryptionInfo XML
            xml = <<~XML
              <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">
                <keyData saltSize="#{key_data_salt.bytesize}" blockSize="#{block_size}" keyBits="#{key_bits}" hashSize="#{hash_size}" cipherAlgorithm="#{cipher_algorithm}" cipherChaining="#{cipher_chaining}" hashAlgorithm="#{hash_algorithm}" saltValue="#{[key_data_salt].pack("m0")}"/>
                <dataIntegrity encryptedHmacKey="#{[encrypted_hmac_key].pack("m0")}" encryptedHmacValue="#{[encrypted_hmac_val].pack("m0")}"/>
                <keyEncryptors>
                  <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
                    <p:encryptedKey spinCount="#{spin_count}" saltSize="#{encryptor_salt.bytesize}" blockSize="#{block_size}" keyBits="#{key_bits}" hashSize="#{hash_size}" cipherAlgorithm="#{cipher_algorithm}" cipherChaining="#{cipher_chaining}" hashAlgorithm="#{hash_algorithm}" saltValue="#{[encryptor_salt].pack("m0")}" encryptedVerifierHashInput="#{[encrypted_verifier_input].pack("m0")}" encryptedVerifierHashValue="#{[encrypted_verifier_val].pack("m0")}" encryptedKeyValue="#{[encrypted_key_value].pack("m0")}"/>
                  </keyEncryptor>
                </keyEncryptors>
              </encryption>
            XML

            # EncryptionInfo stream header: versionMajor=4, versionMinor=4, flags=0x40
            info_stream = [4, 4, 0x40].pack("vvV") + xml.b

            {
              "EncryptionInfo" => info_stream,
              "EncryptedPackage" => encrypted_package_stream
            }
          end

          private

          def parse_encryption_info(bytes)
            raise Xlsxrb::DecryptionError, "EncryptionInfo stream is too short" if bytes.bytesize < 8

            major, minor, = bytes[0, 8].unpack("vvV")
            raise Xlsxrb::DecryptionError, "Unsupported encryption version #{major}.#{minor}" unless major == 4 && minor == 4

            xml_str = bytes[8..] || ""
            begin
              doc = REXML::Document.new(xml_str)
            rescue StandardError => e
              raise Xlsxrb::DecryptionError, "Malformed EncryptionInfo XML: #{e.message}"
            end

            root = doc.root
            raise Xlsxrb::DecryptionError, "Invalid EncryptionInfo XML" unless root

            key_data = root.elements["keyData"]
            key_encryptor = root.elements["keyEncryptors/keyEncryptor/encryptedKey"] || root.elements["keyEncryptors/keyEncryptor/*"]
            raise Xlsxrb::DecryptionError, "Missing keyEncryptor element" unless key_encryptor

            spin_count = key_encryptor.attributes["spinCount"]&.to_i || 100_000
            raise Xlsxrb::DecryptionError, "spinCount out of allowed bounds: #{spin_count}" if spin_count.negative? || spin_count > 10_000_000

            key_attrs = key_data&.attributes
            {
              key_data_salt: key_attrs&.[]("saltValue").to_s.unpack1("m0") || "".b,
              key_data_hash_algorithm: key_attrs&.[]("hashAlgorithm") || "SHA512",
              key_data_cipher_algorithm: key_attrs&.[]("cipherAlgorithm") || "AES",
              key_data_key_bits: key_attrs&.[]("keyBits")&.to_i || 256,
              key_data_block_size: key_attrs&.[]("blockSize")&.to_i || 16,

              spin_count: spin_count,
              encryptor_salt: key_encryptor.attributes["saltValue"].to_s.unpack1("m0") || "".b,
              hash_algorithm: key_encryptor.attributes["hashAlgorithm"] || "SHA512",
              cipher_algorithm: key_encryptor.attributes["cipherAlgorithm"] || "AES",
              key_bits: key_encryptor.attributes["keyBits"]&.to_i || 256,
              hash_size: key_encryptor.attributes["hashSize"]&.to_i || 64,
              encrypted_verifier_hash_input: key_encryptor.attributes["encryptedVerifierHashInput"].to_s.unpack1("m0") || "".b,
              encrypted_verifier_hash_value: key_encryptor.attributes["encryptedVerifierHashValue"].to_s.unpack1("m0") || "".b,
              encrypted_key_value: key_encryptor.attributes["encryptedKeyValue"].to_s.unpack1("m0") || "".b
            }
          end

          def digest_for(name)
            case name.to_s.upcase.gsub(/[^A-Z0-9]/, "")
            when "SHA512" then OpenSSL::Digest.new("SHA512")
            when "SHA384" then OpenSSL::Digest.new("SHA384")
            when "SHA256" then OpenSSL::Digest.new("SHA256")
            when "SHA1"   then OpenSSL::Digest.new("SHA1")
            else
              raise Xlsxrb::DecryptionError, "Unsupported hash algorithm: #{name}"
            end
          end

          def cipher_name_for(algo, key_bits)
            case algo.to_s.upcase
            when "AES"
              "aes-#{key_bits}-cbc"
            else
              raise Xlsxrb::DecryptionError, "Unsupported cipher algorithm: #{algo}"
            end
          end

          def derive_h_final(digest, password, salt, spin_count)
            pw_utf16 = password.encode("UTF-16LE").b
            h = digest.digest(salt + pw_utf16)

            spin_count.times do |i|
              h = digest.digest([i].pack("V") + h)
            end
            h
          end

          def derive_key(digest, h_final, block_key, key_bytes)
            digest.digest(h_final + block_key)[0, key_bytes]
          end

          def aes_decrypt(cipher_name, key, init_vector, data)
            cipher = OpenSSL::Cipher.new(cipher_name)
            cipher.decrypt
            cipher.key = key
            cipher.iv = init_vector[0, 16]
            cipher.padding = 0
            cipher.update(data) + cipher.final
          end

          def aes_encrypt(cipher_name, key, init_vector, data)
            cipher = OpenSSL::Cipher.new(cipher_name)
            cipher.encrypt
            cipher.key = key
            cipher.iv = init_vector[0, 16]
            cipher.padding = 0
            # Ensure data is multiple of 16 bytes
            if (data.bytesize % 16) != 0
              pad_len = 16 - (data.bytesize % 16)
              data += ("\x00".b * pad_len)
            end
            cipher.update(data) + cipher.final
          end

          def aes_decrypt_block(cipher_name, key, init_vector, data)
            cipher = OpenSSL::Cipher.new(cipher_name)
            cipher.decrypt
            cipher.key = key
            cipher.iv = init_vector[0, 16]
            cipher.padding = 0
            cipher.update(data) + cipher.final
          end

          def aes_encrypt_block(cipher_name, key, init_vector, data)
            cipher = OpenSSL::Cipher.new(cipher_name)
            cipher.encrypt
            cipher.key = key
            cipher.iv = init_vector[0, 16]
            cipher.padding = 0
            cipher.update(data) + cipher.final
          end
        end
      end
    end
  end
end
