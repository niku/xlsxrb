# frozen_string_literal: true

require "test_helper"
require_relative "../../../lib/xlsxrb/ooxml/cfb"

module Xlsxrb
  module Ooxml
    class CfbTest < Test::Unit::TestCase
      def test_cfb_write_and_read_streams
        streams = {
          "EncryptionInfo" => "Sample Encryption Info XML data <encryption>...</encryption>".b,
          "EncryptedPackage" => ("A" * 5000).b # Spans multiple sectors (> 512 bytes)
        }

        cfb_bytes = Cfb::Writer.write(streams)
        assert_operator cfb_bytes.bytesize, :>, 512
        assert_true Cfb::Reader.cfb?(cfb_bytes)

        reader = Cfb::Reader.new(cfb_bytes)
        assert_equal %w[EncryptionInfo EncryptedPackage], reader.stream_names

        info_read = reader.read_stream("EncryptionInfo")
        assert_equal streams["EncryptionInfo"], info_read

        pkg_read = reader.read_stream("EncryptedPackage")
        assert_equal streams["EncryptedPackage"], pkg_read
      end

      def test_cfb_non_cfb_data
        plain_zip = "PK\x03\x04SomeZipData".b
        assert_false Cfb::Reader.cfb?(plain_zip)
        assert_raise(Xlsxrb::Error) do
          Cfb::Reader.new(plain_zip)
        end
      end
    end
  end
end
