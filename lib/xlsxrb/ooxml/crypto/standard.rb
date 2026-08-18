# frozen_string_literal: true

# rbs_inline: enabled

require "openssl"
require "securerandom"

module Xlsxrb
  module Ooxml
    module Crypto
      # Implements Microsoft Office Standard Encryption specified in [MS-OFFCRYPTO] Section 2.3.6.
      class Standard
        CSP_NAME = "Microsoft Enhanced RSA and AES Cryptographic Provider\x00".encode("UTF-16LE").b.freeze
        ITERATION_COUNT = 50_000

        class << self
          # Decrypts a Standard-encrypted package.
          def decrypt(encryption_info_bytes, encrypted_package_bytes, password)
            salt, encrypted_verifier, verifier_hash_size, encrypted_verifier_hash = parse_encryption_info(encryption_info_bytes)
            password_str = password.to_s

            # Derive key K
            k = derive_key(password_str, salt)

            # Verify password
            verifier = aes_ecb_decrypt(k, encrypted_verifier)
            expected_hash = OpenSSL::Digest::SHA1.digest(verifier)

            decrypted_hash = aes_ecb_decrypt(k, encrypted_verifier_hash)[0, verifier_hash_size]
            raise Xlsxrb::InvalidPasswordError, "Standard encryption password verification failed" unless OpenSSL.secure_compare(decrypted_hash, expected_hash)

            # Decrypt package stream
            raise Xlsxrb::DecryptionError, "Encrypted package stream is too short" if encrypted_package_bytes.bytesize < 8

            total_size = encrypted_package_bytes[0, 8].unpack1("Q<")
            raise Xlsxrb::DecryptionError, "Encrypted package size is invalid or exceeds limits" if total_size.negative? || total_size > 0x400_000_000

            encrypted_data = encrypted_package_bytes[8..] || "".b

            decrypted = aes_ecb_decrypt(k, encrypted_data)
            decrypted[0, total_size] || "".b
          end

          # Encrypts plain zip data into Standard Encryption streams.
          def encrypt(plain_bytes, password)
            password_str = password.to_s
            salt = SecureRandom.random_bytes(16)
            verifier = SecureRandom.random_bytes(16)

            # Derive key K
            k = derive_key(password_str, salt)

            # Encrypt verifier
            encrypted_verifier = aes_ecb_encrypt(k, verifier)

            # Encrypt verifier hash
            verifier_hash = OpenSSL::Digest::SHA1.digest(verifier)
            padded_hash = verifier_hash.ljust(32, "\x00".b)
            encrypted_verifier_hash = aes_ecb_encrypt(k, padded_hash)

            # Build EncryptionInfo Stream
            header_size = 32 + CSP_NAME.bytesize
            info_stream = +""
            # Version & Flags: vMajor=3, vMinor=2, Flags=0x24 (CryptoAPI AES-128)
            info_stream << [3, 2, 0x24].pack("vvV")
            # EncryptionHeader
            info_stream << [header_size].pack("V")
            info_stream << [0x24, 0, 0x0000660E, 0x00008004, 128, 0x00000018, 0, 0].pack("V8")
            info_stream << CSP_NAME
            # EncryptionVerifier
            info_stream << [16].pack("V") # Salt size
            info_stream << salt
            info_stream << encrypted_verifier # 16 bytes
            info_stream << [20].pack("V") # Verifier hash size (SHA-1 = 20)
            info_stream << encrypted_verifier_hash # 32 bytes

            # Build EncryptedPackage Stream
            padded_plain = plain_bytes
            padded_plain = padded_plain.ljust(padded_plain.bytesize + 16 - (padded_plain.bytesize % 16), "\x00".b) if (padded_plain.bytesize % 16) != 0

            encrypted_pkg_data = aes_ecb_encrypt(k, padded_plain)
            pkg_stream = [plain_bytes.bytesize].pack("Q<") + encrypted_pkg_data

            {
              "EncryptionInfo" => info_stream,
              "EncryptedPackage" => pkg_stream
            }
          end

          private

          def parse_encryption_info(bytes)
            raise Xlsxrb::DecryptionError, "EncryptionInfo stream too short" if bytes.bytesize < 40

            major, minor, = bytes[0, 8].unpack("vvV")
            raise Xlsxrb::DecryptionError, "Unsupported standard encryption version #{major}.#{minor}" unless [2, 3, 4].include?(major) && minor == 2

            header_size = bytes[8, 4].unpack1("V") || 0
            verifier_offset = 12 + header_size
            raise Xlsxrb::DecryptionError, "EncryptionInfo header offset out of bounds" if verifier_offset + 56 > bytes.bytesize

            salt_size = bytes[verifier_offset, 4].unpack1("V") || 0
            raise Xlsxrb::DecryptionError, "Invalid salt size in Standard encryption" if salt_size > 64 || (verifier_offset + 4 + salt_size + 48) > bytes.bytesize

            salt = bytes[verifier_offset + 4, salt_size]
            encrypted_verifier = bytes[verifier_offset + 4 + salt_size, 16]

            hash_size_offset = verifier_offset + 4 + salt_size + 16
            verifier_hash_size = bytes[hash_size_offset, 4].unpack1("V") || 0
            encrypted_verifier_hash = bytes[hash_size_offset + 4, 32]

            [salt, encrypted_verifier, verifier_hash_size, encrypted_verifier_hash]
          end

          def derive_key(password, salt, block_num = 0)
            pw_utf16 = password.encode("UTF-16LE").b
            h = OpenSSL::Digest::SHA1.digest(salt + pw_utf16)

            ITERATION_COUNT.times do |i|
              h = OpenSSL::Digest::SHA1.digest([i].pack("V") + h)
            end

            x = OpenSSL::Digest::SHA1.digest(h + [block_num].pack("V"))
            buf1 = x.ljust(64, "\x00".b).bytes.map { |b| b ^ 0x36 }.pack("C*")
            k1 = OpenSSL::Digest::SHA1.digest(buf1)

            buf2 = x.ljust(64, "\x00".b).bytes.map { |b| b ^ 0x5C }.pack("C*")
            k2 = OpenSSL::Digest::SHA1.digest(buf2)

            (k1 + k2)[0, 16]
          end

          def aes_ecb_encrypt(key, data)
            cipher = OpenSSL::Cipher.new("aes-128-ecb")
            cipher.encrypt
            cipher.key = key
            cipher.padding = 0
            cipher.update(data) + cipher.final
          end

          def aes_ecb_decrypt(key, data)
            cipher = OpenSSL::Cipher.new("aes-128-ecb")
            cipher.decrypt
            cipher.key = key
            cipher.padding = 0
            cipher.update(data) + cipher.final
          end
        end
      end
    end
  end
end
