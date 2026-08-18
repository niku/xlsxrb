# frozen_string_literal: true

# rbs_inline: enabled

require_relative "cfb"
require_relative "crypto/agile"
require_relative "crypto/standard"

module Xlsxrb
  module Ooxml
    # High-level encryption and decryption facade for [MS-OFFCRYPTO] Excel document protection.
    module Crypto
      class << self
        # Checks if the binary data is an encrypted Compound File Binary package.
        def encrypted?(data)
          return false unless Cfb::Reader.cfb?(data)

          begin
            reader = Cfb::Reader.new(data)
            reader.stream_names.any? { |n| n.casecmp?("EncryptionInfo") }
          rescue StandardError
            false
          end
        end

        # Decrypts an encrypted XLSX (CFB) package with the given password.
        def decrypt(cfb_data, password)
          raise Xlsxrb::EncryptedFileError, "Password is required to decrypt this file" if password.nil? || password.to_s.empty?

          reader = Cfb::Reader.new(cfb_data)
          encryption_info = reader.read_stream("EncryptionInfo")
          encrypted_package = reader.read_stream("EncryptedPackage")

          raise Xlsxrb::DecryptionError, "Missing EncryptionInfo or EncryptedPackage stream" if encryption_info.nil? || encrypted_package.nil?

          major, minor = encryption_info[0, 4].unpack("vv")
          if major == 4 && minor == 4
            Agile.decrypt(encryption_info, encrypted_package, password)
          elsif minor == 2
            Standard.decrypt(encryption_info, encrypted_package, password)
          else
            raise Xlsxrb::DecryptionError, "Unsupported encryption version #{major}.#{minor}"
          end
        end

        # Encrypts a plain ZIP payload with the given password into an encrypted CFB package.
        def encrypt(plain_zip_data, password, mode: :standard)
          return plain_zip_data if password.nil? || password.to_s.empty?

          streams = if mode == :agile
                      Agile.encrypt(plain_zip_data, password)
                    else
                      Standard.encrypt(plain_zip_data, password)
                    end
          Cfb::Writer.write(streams)
        end
      end
    end
  end
end
